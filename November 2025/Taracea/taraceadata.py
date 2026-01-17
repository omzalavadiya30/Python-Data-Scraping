import time
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException

# Global variable to store data in case of crash/interrupt
extracted_data = []
output_file = "taracea_full_details.xlsx"

def save_data():
    """Helper function to save current data to Excel"""
    if extracted_data:
        df = pd.DataFrame(extracted_data)
        # Reorder columns for better readability
        cols = ['Category', 'Product URL', 'Product Name', 'SKU', 'Brand', 'Finish', 'Size', 'Description', 
                 'Image 1', 'Image 2', 'Image 3', 'Image 4']
        # Ensure all columns exist
        for col in cols:
            if col not in df.columns:
                df[col] = ""
        
        df = df[cols]
        df.to_excel(output_file, index=False)
        print(f"\n[System] Data saved to {output_file} ({len(df)} records).")

def get_text_safe(driver, xpath):
    """Helper to get text safely without crashing"""
    try:
        return driver.find_element(By.XPATH, xpath).text.strip()
    except:
        return "N/A"

def get_attribute_safe(driver, xpath, attribute):
    """Helper to get attribute safely"""
    try:
        return driver.find_element(By.XPATH, xpath).get_attribute(attribute)
    except:
        return ""

def scrape_taracea_complete():
    # --- Setup Selenium ---
    options = Options()
    # options.add_argument("--headless") 
    options.add_argument("--window-size=1920,1080")
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 10)

    base_url = "https://www.taracea.com"
    
    # List to store initial URL collection
    products_to_visit = [] 

    try:
        print("=== PHASE 1: COLLECTING PRODUCT URLS ===")
        driver.get(base_url)
        time.sleep(3)

        # 1. Open Menu
        print("Locating 'PRODUCTS' menu...")
        products_menu = wait.until(EC.element_to_be_clickable((By.XPATH, "//details[@id='Details-HeaderMenu-1']/summary")))
        products_menu.click()
        time.sleep(1)
        
        # 2. Get Categories
        menu_content = driver.find_element(By.ID, "MegaMenu-Content-1")
        raw_links = menu_content.find_elements(By.XPATH, ".//ul/li/ul/li/a")
        
        categories_map = []
        for link in raw_links:
            c_name = link.get_attribute("innerText").strip()
            c_url = link.get_attribute("href")
            if c_name and "view all" not in c_name.lower():
                categories_map.append({"name": c_name, "url": c_url})

        print(f"Found {len(categories_map)} subcategories.")

        # 3. Pagination & URL Collection
        for cat in categories_map:
            print(f"Collecting URLs from: {cat['name']}")
            driver.get(cat['url'])
            time.sleep(2)
            
            current_page = 1
            cat_urls = set()
            
            while True:
                # Collect URLs on current page
                elements = driver.find_elements(By.XPATH, "//a[contains(@href, '/products/')]")
                for el in elements:
                    u = el.get_attribute("href")
                    if "/products/" in u and u not in cat_urls:
                        cat_urls.add(u)
                        products_to_visit.append({
                            "Category": cat['name'],
                            "Product URL": u
                        })

                # Pagination Logic
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(1)
                next_label = f"Page {current_page + 1}"
                try:
                    next_btn = driver.find_element(By.XPATH, f"//nav[@class='pagination']//a[@aria-label='{next_label}']")
                    driver.execute_script("arguments[0].click();", next_btn)
                    time.sleep(3)
                    current_page += 1
                except:
                    break # End of pagination
            
            print(f"   -> Found {len(cat_urls)} products in {cat['name']}")

        total_products = len(products_to_visit)
        print(f"\n=== PHASE 2: EXTRACTING DETAILS ({total_products} Products) ===")
        
        # --- PHASE 2: DETAIL EXTRACTION ---
        
        for index, item in enumerate(products_to_visit):
            current_count = index + 1
            url = item['Product URL']
            category = item['Category']
            
            try:
                driver.get(url)
                # Wait for h1 to ensure page load
                wait.until(EC.presence_of_element_located((By.ID, "nombreProducto")))
                
                # 1. Product Name (From h1 id="nombreProducto")
                p_name = get_text_safe(driver, "//h1[@id='nombreProducto']")
                
                # Logging Requirement
                print(f"Scraping {current_count}/{total_products} of {category} -> {p_name} -> {url}")

                # 2. SKU (From span id="skupdf")
                # Text often comes as "SKU 96MAR..." so we remove the word "SKU" if needed
                p_sku = get_text_safe(driver, "//span[@id='skupdf']")
                p_sku = p_sku.replace("SKU", "").replace("&nbsp;", "").strip()

                # 3. Brand & Description (Static)
                p_brand = "Taracea"
                p_desc = "N/A"

                # 4. Finish & Size
                # Using the data-option attribute is much more reliable than counting divs
                # We look for the div with data-option="Finish" then find the text inside the Select class
                
                # Finish
                p_finish = "N/A"
                try:
                    finish_el = driver.find_element(By.XPATH, "//div[@data-option='Finish']//div[contains(@class, 'PwzrSelect-select')]")
                    p_finish = finish_el.text.strip()
                except:
                    pass

                # Size
                p_size = "N/A"
                try:
                    size_el = driver.find_element(By.XPATH, "//div[@data-option='Size']//div[contains(@class, 'PwzrSelect-select')]")
                    p_size = size_el.text.strip()
                except:
                    pass

                # 5. Images
                # Extract up to 4 images from the slider list
                images_found = []
                try:
                    img_elements = driver.find_elements(By.XPATH, "//ul[contains(@class, 'product__media-list')]//li//img")
                    for img in img_elements:
                        src = img.get_attribute("src")
                        if src:
                            # Fix protocol if missing
                            if src.startswith("//"):
                                src = "https:" + src
                            # Remove size parameters to get full image if possible (optional, but cleans url)
                            # src = src.split('?')[0] 
                            if src not in images_found:
                                images_found.append(src)
                except:
                    pass

                # Pad images to ensure we have 4 entries
                while len(images_found) < 4:
                    images_found.append("")

                # --- COMPILE RECORD ---
                record = {
                    "Category": category,
                    "Product URL": url,
                    "Product Name": p_name,
                    "SKU": p_sku,
                    "Brand": p_brand,
                    "Finish": p_finish,
                    "Size": p_size,
                    "Description": p_desc,
                    "Image 1": images_found[0],
                    "Image 2": images_found[1],
                    "Image 3": images_found[2],
                    "Image 4": images_found[3]
                }
                
                extracted_data.append(record)

                # SAVE CONDITION: Every 50 products
                if current_count % 50 == 0:
                    print(f"   [Auto-Save] Saving batch of {len(extracted_data)} products...")
                    save_data()

            except Exception as e:
                print(f"   [Error] Failed on {url}: {e}")
                # Even if failed, try to save partial data or continue
                continue

    except KeyboardInterrupt:
        print("\n\n!!! Script Interrupted by User (Ctrl+C) !!!")
        print("Saving collected data before exiting...")
        save_data()
    
    except Exception as e:
        print(f"\n!!! Critical Script Error: {e} !!!")
        save_data()

    finally:
        print("\n=== SCRAPING FINISHED ===")
        save_data()
        driver.quit()

if __name__ == "__main__":
    scrape_taracea_complete()