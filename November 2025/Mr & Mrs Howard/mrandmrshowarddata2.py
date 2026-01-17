import pandas as pd
import time
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException

def get_text_safe(driver, by, identifier):
    """Helper to extract text safely without crashing if element is missing."""
    try:
        element = driver.find_element(by, identifier)
        return element.text.strip()
    except NoSuchElementException:
        return "N/A"

def scrape_microd_fabrics():
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    # options.add_argument('--headless') # Uncomment if you don't want to see the browser
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    # --- PHASE 1: Collect URLs ---
    start_url = "https://mrandmrshoward.microdinc.com/itembrowser.aspx?action=attributes&itemtype=fabric&brand=mr-and-mrs-howard"
    product_urls = []
    
    print("--- PHASE 1: Collecting Product URLs ---")
    try:
        driver.get(start_url)
        wait = WebDriverWait(driver, 20)
        
        # Handle Pagination via "View All"
        try:
            view_all_element = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "pagall")))
            view_all_url = view_all_element.get_attribute("href")
            print("Navigating to View All page...")
            driver.get(view_all_url)
            wait.until(EC.presence_of_element_located((By.CLASS_NAME, "ProductThumbnail")))
            time.sleep(5) 
        except Exception as e:
            print(f"Staying on current page (View All not found/loaded). Error: {e}")

        # Extract Links
        anchors = driver.find_elements(By.TAG_NAME, "a")
        for a in anchors:
            try:
                href = a.get_attribute("href")
                if href and "iteminformation.aspx" in href:
                    if href not in product_urls:
                        product_urls.append(href)
            except:
                continue
        
        print(f"Total Unique Products Found: {len(product_urls)}")

    except Exception as e:
        print(f"Error during URL collection: {e}")
        driver.quit()
        return

    # --- PHASE 2: Extract Product Details ---
    print("\n--- PHASE 2: Extracting Product Details ---")
    
    scraped_data = []
    total_products = len(product_urls)
    
    # Helper function to save data
    def save_data_to_excel():
        if scraped_data:
            df = pd.DataFrame(scraped_data)
            df.to_excel("microd_fabric_details.xlsx", index=False)
            print("\n[SYSTEM] Data saved to 'microd_fabric_details.xlsx'")
        else:
            print("\n[SYSTEM] No data to save.")

    try:
        for index, url in enumerate(product_urls):
            current_count = index + 1
            
            try:
                driver.get(url)
                # Wait strictly for the name to ensure page load
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "ProductInfoSpanValueDescription"))
                )
                
                # 1. Category (Fixed)
                category = "Fabrics"
                
                # 2. Brand (Fixed)
                brand = "Mr. and Mrs. Howard for Sherrill Furniture"
                
                # 3. Name
                name = get_text_safe(driver, By.CLASS_NAME, "ProductInfoSpanValueDescription")
                
                # 4. SKU
                sku = get_text_safe(driver, By.CLASS_NAME, "ProductInfoSpanValueSKU")
                
                # 5. Description (Text)
                description = get_text_safe(driver, By.CLASS_NAME, "ProductDetailsParagraphSEO")
                
                # 6. Full HTML Description
                try:
                    html_desc_element = driver.find_element(By.CLASS_NAME, "ProductInfo")
                    full_html = html_desc_element.get_attribute("outerHTML") # or outerHTML
                except NoSuchElementException:
                    full_html = "N/A"

                # 7. Attributes
                cleaning_code = get_text_safe(driver, By.CLASS_NAME, "ProductInfoSpanValueCleaningCode")
                color = get_text_safe(driver, By.CLASS_NAME, "ProductInfoSpanValueColor")
                color_family = get_text_safe(driver, By.CLASS_NAME, "ProductInfoSpanValueColorFamily")
                cover_collection = get_text_safe(driver, By.CLASS_NAME, "ProductInfoSpanValueCoverCollection")
                cover_type = get_text_safe(driver, By.CLASS_NAME, "ProductInfoSpanValueCoverType")
                fabric_width = get_text_safe(driver, By.CLASS_NAME, "ProductInfoSpanValueFabricWidth")
                direction = get_text_safe(driver, By.CLASS_NAME, "ProductInfoSpanValueDirection")
                repeatheight = get_text_safe(driver, By.CLASS_NAME, "ProductInfoSpanValueRepeatHeight")
                repeatwidth = get_text_safe(driver, By.CLASS_NAME, "ProductInfoSpanValueRepeatWidth")

                # 8. Image
                # Target the A tag with ID 'newItemImageControl_lnkMagicZoom' and get href
                image_url = "N/A"
                try:
                    img_element = driver.find_element(By.ID, "newItemImageControl_lnkMagicZoom")
                    image_url = img_element.get_attribute("href")
                    # Fix protocol if it starts with //
                    if image_url and image_url.startswith("//"):
                        image_url = "https:" + image_url
                except NoSuchElementException:
                    pass

                # Store Data
                product_data = {
                    "Category": category,
                    "Product Url": url,
                    "Product Name": name,
                    "SKU": sku,
                    "Brand": brand,
                    "Cleaning Code": cleaning_code,
                    "Color": color,
                    "Color Family": color_family,
                    "Cover Collection": cover_collection,
                    "Cover Type": cover_type,
                    "Fabric Width": fabric_width,
                    "Direction": direction,
                    "Repeat Height": repeatheight,
                    "Repeat Width": repeatwidth,
                    "Description": description,
                    "Full Description HTML": full_html,
                    "Image Url": image_url,
                }
                
                scraped_data.append(product_data)
                
                # --- LOGGING ---
                # Log format: Scraping current/total -> product name -> sku -> product url
                print(f"Scraping {current_count}/{total_products} -> {name} -> {sku} -> {url}")

                # --- SAVE EVERY 50 PRODUCTS ---
                if current_count % 50 == 0:
                    print(f"[SYSTEM] 50 items processed. Saving checkpoint...")
                    save_data_to_excel()

            except Exception as inner_e:
                print(f"Failed to scrape {url}. Error: {inner_e}")
                # We continue to the next product even if one fails

    except KeyboardInterrupt:
        print("\n\n[SYSTEM] Script interrupted by user (Ctrl+C). Saving captured data...")
    except Exception as e:
        print(f"\n\n[SYSTEM] Script crashed: {e}. Saving captured data...")
    finally:
        # Final Save
        save_data_to_excel()
        driver.quit()
        print("[SYSTEM] Driver closed. Done.")

if __name__ == "__main__":
    scrape_microd_fabrics()