import time
import pandas as pd
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# --- CONFIGURATION ---
OUTPUT_FILE = "anngish_complete_data.xlsx"
SAVE_EVERY_N_ROWS = 50

def setup_driver():
    driver = webdriver.Chrome()
    driver.maximize_window()
    return driver

def safe_get_text(driver, selector):
    """Helper to get text safely, returns 'N/A' if not found."""
    try:
        element = driver.find_element(By.CSS_SELECTOR, selector)
        return element.text.strip()
    except:
        return "N/A"

def extract_dimensions_from_desc(description_text):
    """Parses description to find dimensions line."""
    if description_text == "N/A":
        return "N/A"
    
    # Split by lines and look for "Dimensions"
    lines = description_text.split('\n')
    for line in lines:
        if "Dimensions" in line:
            # Remove "Dimensions:" prefix if present
            return line.replace("Dimensions:", "").replace("Dimensions", "").strip()
    
    return "N/A"

def get_product_details(driver, url):
    """Visits a product page and extracts all specific fields."""
    driver.get(url)
    time.sleep(2) # Wait for page load

    data = {}
    data['Product URL'] = url
    data['Brand'] = "Ann Gish"

    # 1. Breadcrumb
    # Format: Home > Products > Category > Product
    try:
        breadcrumb_items = driver.find_elements(By.CSS_SELECTOR, ".breadcrumb a, .breadcrumb span")
        texts = [b.text.strip() for b in breadcrumb_items if b.text.strip() and b.text.strip() != "›"]
        data['Category'] = " > ".join(texts)
    except:
        data['Category'] = "N/A"

    # 2. Product Name
    data['Product Name'] = safe_get_text(driver, ".product-single__title-text")

    # 3. SKU
    data['SKU'] = safe_get_text(driver, ".product-single__sku")

    # 4. Price
    # Gets plain text like "$2,125.00"
    try:
        price_elem = driver.find_element(By.CSS_SELECTOR, ".product__price .text-black")
        data['Price'] = price_elem.text.strip()
    except:
        data['Price'] = "N/A"

    # 5. Description & Dimensions
    desc_text = safe_get_text(driver, ".description")
    data['Description'] = desc_text
    data['Dimensions'] = extract_dimensions_from_desc(desc_text)

    # 6. Size
    # Looks for label containing "Select Size" or just "Size"
    try:
        # Try finding the label text specifically
        size_label = driver.find_element(By.XPATH, "//label[contains(., 'Select Size') or contains(., 'Size:')]")
        data['Size'] = size_label.text.replace("Select Size:", "").replace("Size:", "").strip()
    except:
        data['Size'] = "N/A"

    # 7. Images (Image1 to Image4)
    images = []
    
    # A. Main Image
    try:
        main_img = driver.find_element(By.CSS_SELECTOR, ".product-single__media img")
        src = main_img.get_attribute("src")
        if src: images.append(src)
    except:
        pass

    # B. Thumbnails
    try:
        thumbnails = driver.find_elements(By.CSS_SELECTOR, ".image-gallery-thumbnails img")
        for thumb in thumbnails:
            src = thumb.get_attribute("src")
            # Optional: Clean URL to get high res if needed (remove _110x110)
            if src: 
                clean_src = src.replace("_110x110_crop_center", "").replace("_110x110", "") 
                if clean_src not in images: # Avoid duplicates
                    images.append(clean_src)
    except:
        pass

    # Fill Image1 - Image4
    for i in range(4):
        key = f"Image{i+1}"
        if i < len(images):
            data[key] = images[i]
        else:
            data[key] = "N/A"

    return data

def main():
    driver = setup_driver()
    base_url = "https://www.anngish.com"
    driver.get(base_url)
    wait = WebDriverWait(driver, 15)
    actions = ActionChains(driver)

    # --- PHASE 1: COLLECT URLS ---
    category_product_map = [] # Stores {category_name: "X", products: [url1, url2...]}
    
    target_menus = ['bedding', 'home-decor', 'collections']
    collections_exclusions = [
        "Embroideries", "Florals & Botanicals", "Geometrics & Abstracts", 
        "Quilting", "Solids", "Textures", "Jacquard", "BY DESIGN:", "‎" 
    ]

    valid_categories = []

    print("=== PHASE 1: COLLECTING PRODUCT URLS ===")
    
    try:
        # 1. Get Categories
        for menu_name in target_menus:
            menu_btn_selector = f"button[aria-controls='SiteNavLabel-{menu_name}']"
            dropdown_id = f"SiteNavLabel-{menu_name}"
            try:
                menu_btn = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, menu_btn_selector)))
                actions.move_to_element(menu_btn).perform()
                time.sleep(1.5)
                
                dropdown = driver.find_element(By.ID, dropdown_id)
                links = dropdown.find_elements(By.TAG_NAME, "a")
                
                for link in links:
                    name = link.get_attribute("innerText").strip()
                    url = link.get_attribute("href")
                    
                    if not name or not url: continue
                    if menu_name == 'collections' and name in collections_exclusions: continue
                    
                    valid_categories.append({"name": name, "url": url})
            except Exception as e:
                print(f"Menu error {menu_name}: {e}")

        print(f"Total categories to scan: {len(valid_categories)}")

        # 2. Collect URLs per Category (Row-by-Row Scroll)
        for cat in valid_categories:
            print(f"Scanning Category: {cat['name']}...")
            driver.get(cat['url'])
            time.sleep(3)
            
            # Scroll logic
            last_height = driver.execute_script("return document.body.scrollHeight")
            curr_pos = 0
            while True:
                curr_pos += 600
                driver.execute_script(f"window.scrollTo(0, {curr_pos});")
                time.sleep(1) # Fast scroll
                new_height = driver.execute_script("return document.body.scrollHeight")
                if curr_pos >= new_height:
                    time.sleep(2)
                    if driver.execute_script("return document.body.scrollHeight") == new_height:
                        break
                    else:
                        new_height = driver.execute_script("return document.body.scrollHeight")
                last_height = new_height

            # Grab URLs
            try:
                # Specific selector for collection grid
                elems = driver.find_elements(By.CSS_SELECTOR, "div.collection-products .collection-product a.relative")
            except:
                elems = driver.find_elements(By.CSS_SELECTOR, ".collection-product a.relative")
                
            cat_urls = []
            for el in elems:
                u = el.get_attribute("href")
                if u and u not in cat_urls:
                    cat_urls.append(u)
            
            category_product_map.append({
                "category": cat['name'],
                "urls": cat_urls
            })
            print(f"   -> Found {len(cat_urls)} products.")

    except Exception as e:
        print(f"Error in Phase 1: {e}")

    # --- PHASE 2: EXTRACT DETAILS ---
    print("\n=== PHASE 2: EXTRACTING DETAILS ===")
    
    all_extracted_data = []
    total_processed = 0

    try:
        for cat_data in category_product_map:
            cat_name = cat_data['category']
            urls = cat_data['urls']
            total_in_cat = len(urls)
            
            print(f"\n--- Processing Category: {cat_name} ({total_in_cat} items) ---")

            for idx, product_url in enumerate(urls):
                try:
                    # Log scraping progress
                    print(f"Scraping [{idx+1}/{total_in_cat}]: ", end="")
                    
                    details = get_product_details(driver, product_url)
                    
                    # Ensure category name is consistent
                    # details['Source Category'] = cat_name 
                    
                    print(f"{details['Product Name']} -> SKU: {details['SKU']}")
                    
                    all_extracted_data.append(details)
                    total_processed += 1

                    # Save every 50 items
                    if total_processed % SAVE_EVERY_N_ROWS == 0:
                        print(f"...Auto-saving {total_processed} items...")
                        df = pd.DataFrame(all_extracted_data)
                        df.to_excel(OUTPUT_FILE, index=False)

                except KeyboardInterrupt:
                    raise # Allow outer block to catch manual stop
                except Exception as e:
                    print(f"Error on {product_url}: {e}")
                    # Still add partial data with error note? 
                    # For now just skip or store URL with error
                    all_extracted_data.append({
                        # "Source Category": cat_name, 
                        "Product URL": product_url, 
                        "Product Name": "ERROR", 
                        "SKU": "ERROR"
                    })

    except KeyboardInterrupt:
        print("\n!!! Script stopped by user (Ctrl+C) !!!")
    except Exception as e:
        print(f"\n!!! Critical Error: {e} !!!")
    finally:
        # Final Save
        print(f"\nSaving final data ({len(all_extracted_data)} records) to {OUTPUT_FILE}...")
        if all_extracted_data:
            df = pd.DataFrame(all_extracted_data)
            # Reorder columns slightly for better readability
            cols = [
                "Category", "Product URL", "Product Name", "SKU", "Price", 
                "Size", "Dimensions", "Brand", "Description", 
                "Image1", "Image2", "Image3", "Image4"
            ]
            # Only select cols that exist in data
            final_cols = [c for c in cols if c in df.columns]
            df = df[final_cols]
            
            df.to_excel(OUTPUT_FILE, index=False)
            print("Save Complete.")
        
        driver.quit()

if __name__ == "__main__":
    main()