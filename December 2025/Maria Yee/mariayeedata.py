import time
import pandas as pd
import sys
import os
from urllib.parse import urlparse
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Configuration
EXCEL_FILENAME = "mariayee_products.xlsx"
SAVE_INTERVAL = 50

def clean_url(url):
    """Removes query parameters to get the base product URL."""
    if not url: return None
    parsed = urlparse(url)
    return f"{parsed.scheme}://{parsed.netloc}{parsed.path}"

def save_data(data, filename):
    """Helper function to save list of dictionaries to Excel."""
    if not data:
        return
    try:
        df = pd.DataFrame(data)
        # Reorder columns to put Images at the end if they exist
        cols = list(df.columns)
        image_cols = [c for c in cols if c.startswith('Image')]
        base_cols = [c for c in cols if c not in image_cols]
        df = df[base_cols + image_cols]
        
        df.to_excel(filename, index=False)
        print(f"\n[System] Data saved to '{filename}' ({len(df)} records).")
    except Exception as e:
        print(f"[Error] Could not save data: {e}")

def extract_product_details(driver, category_name, product_url):
    """Extracts all details from a single product page."""
    details = {
        "Category": "N/A",
        "Product URL": product_url,
        "Product Name": "N/A",
        "SKU": "N/A",
        "Price": "N/A",
        "Brand": "Maria Yee",
        "Description": "N/A",
        "Dimensions": "N/A",
    }

    try:
        driver.get(product_url)
        wait = WebDriverWait(driver, 5) # Short wait for elements

        # 1. Category Path from URL
        # Format: Category > Products > Product Name (slug)
        try:
            parsed = urlparse(product_url)
            path_parts = parsed.path.strip('/').split('/')
            # Usually /collections/collection-name/products/product-handle
            slug = path_parts[-1].replace('-', ' ').title() if path_parts else "Unknown"
            details["Category"] = f"{category_name} > Products > {slug}"
        except:
            details["Category"] = f"{category_name} > Products"

        # 2. Product Name
        try:
            # <h1 class="product__title..."><span data-zoom-caption="">Name</span></h1>
            title_elem = driver.find_element(By.CSS_SELECTOR, "h1.product__title")
            details["Product Name"] = title_elem.text.strip()
        except:
            details["Product Name"] = "N/A"

        # 3. Price
        try:
            # <div class="product__price" ...><span data-product-price="">$3,300.00</span>
            price_elem = driver.find_element(By.CSS_SELECTOR, "[data-product-price]")
            details["Price"] = price_elem.text.strip()
        except:
            details["Price"] = "N/A"

        # 4. Description (Tab 0)
        try:
            # <div class="tab-content tab-content-0 current rte">
            desc_elem = driver.find_element(By.CSS_SELECTOR, ".tab-content-0")
            details["Description"] = desc_elem.text.strip()
        except:
            pass

        # 5. Dimensions (Tab 1 - Details)
        # We need to click the "Details" tab header if content isn't visible, 
        # but usually getting 'textContent' works even if hidden. 
        # If empty, we try clicking.
        try:
            dim_elem = driver.find_element(By.CSS_SELECTOR, ".tab-content-1")
            dim_text = dim_elem.get_attribute("textContent").strip()
            
            if not dim_text:
                # Try clicking the tab header
                # Look for a list item or link containing "Details"
                try:
                    tab_header = driver.find_element(By.XPATH, "//ul[contains(@class,'tabs')]//li[contains(., 'Details')]")
                    tab_header.click()
                    time.sleep(0.5)
                    dim_text = dim_elem.text.strip()
                except:
                    pass
            
            details["Dimensions"] = dim_text
        except:
            pass

        # 6. Images
        # <div class="product__slide ... data-image-src="...">
        try:
            # Find all slides
            slides = driver.find_elements(By.CSS_SELECTOR, ".product__slide")
            img_count = 1
            for slide in slides:
                # Extract URL from data-image-src or img tag inside
                img_url = slide.get_attribute("data-image-src")
                
                # Fallback to img tag src if data attribute missing
                if not img_url:
                    try:
                        img_tag = slide.find_element(By.TAG_NAME, "img")
                        img_url = img_tag.get_attribute("src")
                        # Clean up query params from Shopify images usually
                        if img_url: 
                            img_url = clean_url(img_url)
                    except:
                        continue

                if img_url:
                    # Fix protocol relative URLs
                    if img_url.startswith("//"):
                        img_url = "https:" + img_url
                    
                    details[f"Image{img_count}"] = img_url
                    img_count += 1
                    
                    if img_count > 10: # Limit to 10 images to prevent massive headers
                        break
        except:
            pass

    except Exception as e:
        print(f"[Warning] Error extracting details for {product_url}: {e}")

    return details

def scrape_mariayee():
    # 1. Setup Chrome Driver
    options = webdriver.ChromeOptions()
    # options.add_argument("--headless") 
    options.add_argument("--start-maximized")
    options.add_argument("--log-level=3") # Suppress console logs
    driver = webdriver.Chrome(options=options)
    
    base_url = "https://www.mariayee.com"
    all_product_data = [] # Stores final detailed data
    
    # Load existing data if file exists to avoid duplicates (Optional logic)
    # if os.path.exists(EXCEL_FILENAME):
    #    existing_df = pd.read_excel(EXCEL_FILENAME)
    #    all_product_data = existing_df.to_dict('records')

    try:
        print(f"Navigating to {base_url}...")
        driver.get(base_url)
        wait = WebDriverWait(driver, 10)

        # --- STEP 1: Get Categories ---
        print("Locating 'Shop' menu...")
        shop_menu_element = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//div[contains(@class, 'menu__item') and .//span[contains(text(), 'Shop')]]")
        ))

        actions = ActionChains(driver)
        actions.move_to_element(shop_menu_element).perform()
        time.sleep(2) 

        category_elements = driver.find_elements(By.CSS_SELECTOR, ".header__dropdown .navlink--child, .header__dropdown .navlink--grandchild")

        categories_to_scrape = []
        seen_urls = set()
        for cat in category_elements:
            name = cat.text.strip()
            url = cat.get_attribute("href")
            
            if url and "collections" in url:
                clean = clean_url(url)
                if clean not in seen_urls:
                    categories_to_scrape.append({"name": name, "url": clean})
                    seen_urls.add(clean)

        print(f"Found {len(categories_to_scrape)} categories.")

        # --- STEP 2: Scrape Each Category ---
        for category in categories_to_scrape:
            cat_name = category['name']
            cat_url = category['url']
            
            print(f"\n--- Processing Category: {cat_name} ---")
            driver.get(cat_url)
            time.sleep(2)

            # A. Collect ALL Product URLs for this category first (Handle Pagination)
            category_product_urls = set()
            
            while True:
                try:
                    # Get product links on current page
                    product_links = driver.find_elements(By.CSS_SELECTOR, "a.product-link")
                    for link in product_links:
                        raw_url = link.get_attribute("href")
                        if raw_url:
                            category_product_urls.add(clean_url(raw_url))
                    
                    # Check for Next Button
                    next_button = driver.find_elements(By.CSS_SELECTOR, "a.pagination-custom__next")
                    if next_button and next_button[0].is_displayed():
                        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", next_button[0])
                        time.sleep(1)
                        next_button[0].click()
                        time.sleep(3)
                    else:
                        break # No more pages
                except Exception as e:
                    print(f"Pagination error: {e}")
                    break
            
            total_products = len(category_product_urls)
            print(f"Found {total_products} unique products in {cat_name}. Starting extraction...")

            # B. Extract Details for each Product
            for i, p_url in enumerate(category_product_urls, 1):
                try:
                    # Extract Data
                    p_data = extract_product_details(driver, cat_name, p_url)
                    all_product_data.append(p_data)

                    # Log: Scrapping current product/ total products of category -> product name -> product url
                    print(f"Scraping {i}/{total_products} of {cat_name} -> {p_data.get('Product Name', 'N/A')} -> {p_url}")

                    # Periodic Save
                    if len(all_product_data) % SAVE_INTERVAL == 0:
                        save_data(all_product_data, EXCEL_FILENAME)

                except KeyboardInterrupt:
                    raise # Re-raise to be caught by outer block
                except Exception as e:
                    print(f"[Error] Failed product {p_url}: {e}")

    except KeyboardInterrupt:
        print("\n[User Interrupt] Script stopped by user. Saving collected data...")
    except Exception as main_e:
        print(f"\n[Critical Error] Script crashed: {main_e}")
    finally:
        # Final Save
        if all_product_data:
            print("\nSaving final data...")
            save_data(all_product_data, EXCEL_FILENAME)
        driver.quit()
        print("Script finished.")

if __name__ == "__main__":
    scrape_mariayee()