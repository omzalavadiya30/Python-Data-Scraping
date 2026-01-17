import time
import pandas as pd
import signal
import sys
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException

# --- Configuration ---
START_URL = "https://sunpan.com"
OUTPUT_FILE = "sunpan_final_data_split.xlsx"
SAVE_EVERY_N = 50

# Global list to store data (for Ctrl+C safety)
all_scraped_data = []

# Exact columns requested (With Split Imperial/Metric)
FINAL_COLUMNS = [
    "Category", "Product url", "Product name", "SKU", "Brand", "Description",
    "Material", "Cover", "Base / Legs", "Base Finish", "Bed Size",
    "Material Finish", "Shape", "Material Content", "Warranty",
    "Assembly Required", "Additional Features", "CARB Compliant",
    "Contract Viable", "Frame", "Outdoor",
    
    # Split Dimensions & Weights
    "Overall Dimensions (Imperial)", "Overall Dimensions (Metric)",
    "Gross Weight (Imperial)", "Gross Weight (Metric)",
    "Net Weight (Imperial)", "Net Weight (Metric)",
    "Carton Weight (Imperial)", "Carton Weight (Metric)",
    "Carton Size (Imperial)", "Carton Size (Metric)",
    "Weight Capacity (Imperial)", "Weight Capacity (Metric)",
    
    # Standard Fields
    "Minimum Order (pcs)", "Pack",
    "Zone", "Origin", "UPC", "CBM", "20FT Container", "40GP Container", "40HQ Container",
    "Image1", "Image2", "Image3", "Image4"
]

# Fields that require splitting (used for logic check)
SPLIT_FIELDS = [
    "Overall Dimensions", "Gross Weight", "Net Weight", 
    "Carton Weight", "Carton Size", "Weight Capacity"
]

def setup_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    # options.add_argument("--headless") 
    driver = webdriver.Chrome(options=options)
    return driver

def save_data_to_excel(data, filename):
    if not data:
        return
    
    df = pd.DataFrame(data)
    
    # Ensure all required columns exist
    for col in FINAL_COLUMNS:
        if col not in df.columns:
            df[col] = "N/A"
            
    # Filter and reorder
    df_final = df[FINAL_COLUMNS]
    
    try:
        df_final.to_excel(filename, index=False)
        print(f"\n[System] Successfully saved {len(df_final)} products to {filename}.")
    except Exception as e:
        print(f"[System] Error saving file: {e}")

def signal_handler(sig, frame):
    print("\n\n[System] Stopping script... Saving collected data...")
    save_data_to_excel(all_scraped_data, OUTPUT_FILE)
    sys.exit(0)

signal.signal(signal.SIGINT, signal_handler)

def scroll_to_element(driver, element):
    try:
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element)
        time.sleep(1)
    except:
        pass

def get_text_content(element):
    try:
        return element.get_attribute("textContent").strip()
    except:
        return ""

def get_text_safe(element, selector):
    try:
        el = element.find_element(By.CSS_SELECTOR, selector)
        return el.get_attribute("textContent").strip()
    except:
        return ""

# ---------------------------------------------------------
# STEP 1: COLLECT CATEGORIES
# ---------------------------------------------------------
def collect_categories(driver):
    print("--- Step 1: Scanning Menu for Categories ---")
    driver.get(START_URL)
    time.sleep(4)
    
    categories = []
    try:
        nav_container = driver.find_element(By.ID, "nav--guest-menu")
        menu_items = nav_container.find_elements(By.TAG_NAME, "details")
        
        for item in menu_items:
            try:
                summary = item.find_element(By.TAG_NAME, "summary")
                main_menu_name = get_text_content(summary)
                
                if not main_menu_name or main_menu_name.lower() in ["register", "search"]: 
                    continue
                
                driver.execute_script("arguments[0].click();", summary)
                time.sleep(1)
                
                sub_links = item.find_elements(By.CSS_SELECTOR, "a.mega-menu__link, a.mega-menu_link--level-2")
                
                for link in sub_links:
                    cat_name = get_text_content(link)
                    cat_url = link.get_attribute("href")
                    
                    if cat_name and cat_url and "products" not in cat_url:
                        if not any(d['url'] == cat_url for d in categories):
                            categories.append({
                                "Main Menu": main_menu_name,
                                "Category Name": cat_name,
                                "url": cat_url
                            })
                
                driver.execute_script("arguments[0].click();", summary)
                time.sleep(0.5)
            except:
                continue
    except Exception as e:
        print(f"Error collecting categories: {e}")
        
    print(f"Total Categories Found: {len(categories)}")
    return categories

# ---------------------------------------------------------
# STEP 2: COLLECT PRODUCT URLS
# ---------------------------------------------------------
def get_product_urls_for_category(driver, category_url):
    driver.get(category_url)
    time.sleep(3)
    
    product_urls = set()
    page = 1
    
    while True:
        products = driver.find_elements(By.CSS_SELECTOR, "a.full-unstyled-link.main-link")
        
        for p in products:
            u = p.get_attribute("href")
            if u: product_urls.add(u)
            
        print(f"   Page {page}: Found {len(products)} products (Total accumulated: {len(product_urls)})")
        
        try:
            next_btn = driver.find_element(By.CSS_SELECTOR, "a[aria-label='Next page']")
            if next_btn.get_attribute("aria-disabled") == "true":
                break
            
            scroll_to_element(driver, next_btn)
            try:
                next_btn.click()
            except ElementClickInterceptedException:
                driver.execute_script("arguments[0].click();", next_btn)
                
            page += 1
            time.sleep(4)
        except NoSuchElementException:
            break
        except Exception:
            break
            
    return list(product_urls)

# ---------------------------------------------------------
# STEP 3: SCRAPE DETAILS
# ---------------------------------------------------------
def scrape_details(driver, product_url, main_menu, category_name):
    driver.get(product_url)
    time.sleep(2)
    
    # Initialize with N/A
    item = {col: "N/A" for col in FINAL_COLUMNS}
    
    item["Product url"] = product_url
    item["Brand"] = "Sunpan"
    
    # 1. Product Name
    try:
        try:
            item["Product name"] = get_text_content(driver.find_element(By.CSS_SELECTOR, "div.product__title h1"))
        except:
            item["Product name"] = get_text_content(driver.find_element(By.CSS_SELECTOR, "div.product__title h2"))
    except:
        pass

    # 2. Category
    item["Category"] = f"Home > {category_name} > {item['Product name']}"

    # 3. SKU
    try:
        sku_txt = get_text_content(driver.find_element(By.CSS_SELECTOR, "p.product__sku"))
        item["SKU"] = sku_txt.replace("SKU:", "").strip()
    except:
        pass

    # 4. Description
    try:
        desc_box = driver.find_element(By.CSS_SELECTOR, ".tab-details__container")
        scroll_to_element(driver, desc_box)
        main_desc = get_text_safe(desc_box, ".rte")
        bullets = [get_text_content(li) for li in desc_box.find_elements(By.TAG_NAME, "li")]
        full_desc = main_desc
        if bullets:
            full_desc += " " + " ".join(bullets)
        item["Description"] = full_desc.strip()
    except:
        pass

    # 5. Specs Table
    try:
        const_table = driver.find_element(By.CSS_SELECTOR, ".tab-construction__table")
        rows = const_table.find_elements(By.TAG_NAME, "tr")
        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) == 2:
                key = get_text_content(cols[0])
                val = get_text_content(cols[1])
                if key in item:
                    item[key] = val
    except:
        pass

    # 6. Dimensions & Weight (SPLIT LOGIC)
    try:
        dim_table = driver.find_element(By.CSS_SELECTOR, ".tab-dimensions__table")
        rows = dim_table.find_elements(By.TAG_NAME, "tr")
        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) == 2:
                header = get_text_content(cols[0])
                
                # Logic for Splitting Imperial/Metric fields
                if header in SPLIT_FIELDS:
                    try:
                        # Extract hidden text content for separate units
                        imp = cols[1].find_element(By.CSS_SELECTOR, ".imperial").get_attribute("textContent").strip()
                        met = cols[1].find_element(By.CSS_SELECTOR, ".metric").get_attribute("textContent").strip()
                        
                        item[f"{header} (Imperial)"] = imp
                        item[f"{header} (Metric)"] = met
                    except:
                        # Fallback: Put same value in both or just Imperial if metric missing
                        val = get_text_content(cols[1])
                        item[f"{header} (Imperial)"] = val
                        item[f"{header} (Metric)"] = val
                
                # Logic for Non-Split fields (Pack, Minimum Order)
                elif header in item:
                    item[header] = get_text_content(cols[1])
    except:
        pass

    # 7. Shipping Info
    try:
        ship_table = driver.find_element(By.CSS_SELECTOR, ".tab-shipping__table")
        rows = ship_table.find_elements(By.TAG_NAME, "tr")
        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) == 2:
                key = get_text_content(cols[0])
                val = get_text_content(cols[1])
                if key in item:
                    item[key] = val
    except:
        pass

    # 8. Images
    try:
        # We use a CSS selector to find the list, ignoring the dynamic ID number.
        # We target the <ul> with class 'product__media-list' and finding <img> tags inside 'li' items.
        images_found = []
        
        # This finds all images inside the slider list
        img_elements = driver.find_elements(By.CSS_SELECTOR, "ul.product__media-list li.product__media-item img")
        
        for img in img_elements:
            if len(images_found) >= 4:
                break
                
            try:
                src = img.get_attribute("src")
                
                # If src is empty or a placeholder, try srcset
                if not src or "data:image" in src:
                    srcset = img.get_attribute("srcset")
                    if srcset:
                        # Grab the last URL in the srcset (usually the highest resolution)
                        src = srcset.split(",")[-1].strip().split(" ")[0]

                if src:
                    # 1. Handle Protocol (add https: if missing)
                    if src.startswith("//"):
                        src = "https:" + src
                        
                    # 2. Clean URL (Remove ?v=... and &width=... to get the original file)
                    # Example input: https://sunpan.com/files/102937.jpg?v=1765959244&width=1946
                    # Result: https://sunpan.com/files/102937.jpg
                    clean_url = src.split("?")[0]
                    
                    # 3. Validation & Deduplication
                    # Check if it's a valid image extension and not already added
                    if clean_url not in images_found and (".jpg" in clean_url.lower() or ".png" in clean_url.lower() or ".jpeg" in clean_url.lower()):
                        images_found.append(clean_url)
            except:
                continue
        
        # Assign found images to the item dictionary
        for i, img_url in enumerate(images_found):
            item[f"Image{i+1}"] = img_url
            
    except Exception as e:
        print(f"Error extracting images: {e}")

    return item

# ---------------------------------------------------------
# MAIN
# ---------------------------------------------------------
def main():
    driver = setup_driver()
    
    try:
        categories = collect_categories(driver)
        
        for cat_idx, cat in enumerate(categories):
            cat_name = cat['Category Name']
            main_menu = cat['Main Menu']
            cat_url = cat['url']
            
            print(f"\n==================================================")
            print(f"Processing Category {cat_idx+1}/{len(categories)}: {cat_name}")
            print(f"URL: {cat_url}")
            print(f"==================================================")
            
            product_links = get_product_urls_for_category(driver, cat_url)
            total = len(product_links)
            
            if total == 0:
                print("No products found.")
                continue
                
            for prod_idx, link in enumerate(product_links):
                try:
                    product_data = scrape_details(driver, link, main_menu, cat_name)
                    all_scraped_data.append(product_data)
                    
                    print(f"Scraping {prod_idx+1}/{total} of {cat_name} -> {product_data['Product name']} -> {product_data['SKU']} -> {link}")
                    
                    if len(all_scraped_data) % SAVE_EVERY_N == 0:
                        save_data_to_excel(all_scraped_data, OUTPUT_FILE)
                        
                except Exception as e:
                    print(f"Error scraping product {link}: {e}")
                    
    except Exception as e:
        print(f"Global Error: {e}")
        
    finally:
        driver.quit()
        save_data_to_excel(all_scraped_data, OUTPUT_FILE)
        print("\n--- Script Finished ---")

if __name__ == "__main__":
    main()