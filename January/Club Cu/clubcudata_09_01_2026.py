import time
import pandas as pd
import sys
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, 
    NoSuchElementException, 
    ElementClickInterceptedException
)

# --- CONFIGURATION ---
BASE_URL = "https://shop.cohab.space/"
OUTPUT_FILE = "cohab_space_full_inventory.xlsx"
SAVE_INTERVAL = 50

def setup_driver():
    """Initializes the browser."""
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    driver = webdriver.Chrome(options=options)
    return driver

def clean_text(text):
    """Helper to clean whitespace. Returns None if empty."""
    if text:
        cleaned = text.replace('\n', ' ').strip()
        return cleaned if cleaned else None
    return None

def get_high_res_image(url):
    """Removes dimensions like -150x150 from URL to get full size image."""
    if not url: return "N/A"
    clean_url = re.sub(r'-\d+x\d+(?=\.\w+$)', '', url)
    return clean_url

def validate_data(data):
    """Ensures no field is blank/empty. Sets 'N/A' if missing."""
    for key in data:
        if data[key] is None or data[key] == "" or str(data[key]).strip() == "":
            data[key] = "N/A"
        
        # specific fix for the "Select a Variation" issue
        if key == "SKU" and "Select" in str(data[key]):
            data[key] = "N/A"
            
    return data

# --- STEP 1: CATEGORY COLLECTION ---
def get_categories(driver):
    print("--- Step 1: Collecting Categories ---")
    driver.get(BASE_URL)
    time.sleep(5) 

    unique_categories = {} 
    
    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, "cohab-category-list"))
        )
        
        slide_elements = driver.find_elements(By.XPATH, "//div[contains(@class, 'cohab-category-item')]//a")
        
        for link_el in slide_elements:
            try:
                url = link_el.get_attribute("href")
                title_element = link_el.find_element(By.XPATH, ".//div[contains(@class, 'cohab-category-title')]//h3")
                name = title_element.get_attribute("textContent").strip()
                
                if url and name:
                    if url not in unique_categories:
                        unique_categories[url] = name
            except Exception:
                continue
                
    except TimeoutException:
        print("Error: Could not find category slider.")
    
    category_list = [{"Name": name, "URL": url} for url, name in unique_categories.items()]
    print(f"-> Found {len(category_list)} unique categories.")
    return category_list

# --- STEP 2: LOAD MORE HANDLER ---
def handle_load_more(driver):
    print("   > Checking for 'Load More' buttons...")
    while True:
        try:
            load_more_btn = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'e-loop__load-more')]//a"))
            )
            
            if not load_more_btn.is_displayed():
                break

            driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", load_more_btn)
            time.sleep(2) 

            driver.execute_script("arguments[0].click();", load_more_btn)
            
            try:
                WebDriverWait(driver, 2).until(
                    EC.visibility_of_element_located((By.CLASS_NAME, "e-load-more-spinner"))
                )
                WebDriverWait(driver, 15).until(
                    EC.invisibility_of_element_located((By.CLASS_NAME, "e-load-more-spinner"))
                )
            except TimeoutException:
                time.sleep(3)

        except (TimeoutException, NoSuchElementException):
            print("   > 'Load More' button gone. All products loaded.")
            break
        except Exception:
            break

# --- STEP 3: COLLECT PRODUCT URLS ---
def get_product_urls_from_category(driver, category_url):
    driver.get(category_url)
    time.sleep(4)
    handle_load_more(driver)

    product_urls = []
    products = driver.find_elements(By.XPATH, "//div[contains(@class, 'e-loop-item')]//a[contains(@class, 'elementor-element')]")
    
    for prod in products:
        url = prod.get_attribute("href")
        if url and url not in product_urls:
            product_urls.append(url)
            
    return product_urls

# --- STEP 4: SCRAPE PRODUCT DETAILS ---
def scrape_product_details(driver, product_url, category_name):
    driver.get(product_url)
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )
    except TimeoutException:
        print(f"   > Timeout loading {product_url}")
        return None

    data = {
        "Category": None,
        "Product URL": product_url,
        "Product Name": None,
        "SKU": None,
        "Price": None,
        "Brand": "Cohab Space",
        "Description": None,
        "Dimensions": None,
        "Image1": None, "Image2": None, "Image3": None, "Image4": None
    }

    # 1. Category Path
    try:
        breadcrumb = driver.find_element(By.CLASS_NAME, "woocommerce-breadcrumb").text
        data["Category"] = breadcrumb.replace('\n', ' / ')
    except: pass

    # 2. Product Name
    try:
        data["Product Name"] = driver.find_element(By.CLASS_NAME, "product_title").text.strip()
    except: pass

    # --- SKU & DIMENSIONS EXTRACTION ---
    
    # Strategy A: Check for Grouped Product Table First (The fix for your issue)
    # matches <tr class="...woocommerce-grouped-product-list-item...">
    try:
        grouped_rows = driver.find_elements(By.XPATH, "//tr[contains(@class, 'woocommerce-grouped-product-list-item')]")
        if grouped_rows:
            # "Extract from 1 data" -> Get the first row
            first_row = grouped_rows[0]
            
            # Extract SKU from this row
            try:
                # matches <div class="sku">CH-TKS241009-B</div>
                sku_div = first_row.find_element(By.CLASS_NAME, "sku")
                data["SKU"] = sku_div.get_attribute("textContent").strip()
            except: pass

            # Extract Dimensions from this row
            try:
                # matches <p class="dimensions">...</p>
                dim_p = first_row.find_element(By.CLASS_NAME, "dimensions")
                raw_dim = dim_p.get_attribute("textContent").strip()
                data["Dimensions"] = raw_dim.replace("Dimensions (in):", "").replace("Dimensions:", "").strip()
            except: pass
    except: pass

    # Strategy B: Standard Single Product Fallback (if Grouped strategy didn't find data)
    
    # SKU Fallback
    if not data["SKU"] or "Select" in str(data["SKU"]): # Check if invalid
        data["SKU"] = None # Reset if invalid
        try:
            # "Code:" text
            sku_elem = driver.find_element(By.XPATH, "//p[contains(text(), 'Code:')]")
            clean_sku = sku_elem.text.replace("Code:", "").strip()
            if "Select" not in clean_sku:
                 data["SKU"] = clean_sku
        except: pass
        
        if not data["SKU"]:
            try:
                # standard .sku class
                sku_elem = driver.find_element(By.CSS_SELECTOR, ".product_meta .sku")
                data["SKU"] = sku_elem.text.strip()
            except: pass

    # Dimensions Fallback
    if not data["Dimensions"]:
        try:
            dim_elem = driver.find_element(By.CLASS_NAME, "product-dimensions")
            data["Dimensions"] = dim_elem.text.replace("Dimensions:", "").strip()
        except: pass
    
    if not data["Dimensions"]:
        try:
             # Hover <p class="dimensions"> in main area
            dim_elems = driver.find_elements(By.XPATH, "//p[contains(@class, 'dimensions')]")
            if dim_elems:
                raw_dim = dim_elems[0].get_attribute("textContent").strip()
                data["Dimensions"] = raw_dim.replace("Dimensions (in):", "").replace("Dimensions:", "").strip()
        except: pass

    # 4. Price
    try:
        price_elem = driver.find_element(By.CSS_SELECTOR, ".woocommerce-Price-amount bdi")
        data["Price"] = price_elem.text.strip()
    except: pass

    # 5. Description
    full_desc = ""
    try:
        short_desc = driver.find_element(By.CLASS_NAME, "woocommerce-product-details__short-description").text
        full_desc += short_desc + "\n"
    except: pass
    try:
        long_desc = driver.find_element(By.XPATH, "//div[contains(@class, 'elementor-widget-text-editor')]//div[@class='elementor-widget-container']").text
        full_desc += long_desc
    except: pass
    data["Description"] = clean_text(full_desc)

    # 7. Images
    images = []
    # A. Main Image
    try:
        main_img = driver.find_element(By.XPATH, "//div[contains(@class, 'swiper-slide-active')]//a")
        img_url = main_img.get_attribute("href")
        if not img_url:
            img_src = main_img.find_element(By.TAG_NAME, "img").get_attribute("src")
            img_url = get_high_res_image(img_src)
        if img_url:
            images.append(img_url)
    except: pass
    # B. Thumbnails
    try:
        thumbs = driver.find_elements(By.XPATH, "//div[contains(@class, 'thumbnail-gallery')]//div[contains(@class, 'swiper-slide')]//img")
        for thumb in thumbs:
            src = thumb.get_attribute("src")
            high_res = get_high_res_image(src)
            if high_res and high_res not in images:
                images.append(high_res)
    except: pass

    for i, img in enumerate(images[:4]):
        data[f"Image{i+1}"] = img

    # FINAL VALIDATION
    data = validate_data(data)
    
    return data

# --- MAIN EXECUTION ---
def main():
    driver = setup_driver()
    all_extracted_data = []

    try:
        # Step 1
        categories = get_categories(driver)

        # Step 2
        for cat in categories:
            cat_name = cat['Name']
            print(f"\n--- Processing Category: {cat_name} ---")
            
            product_urls = get_product_urls_from_category(driver, cat['URL'])
            print(f"   > Found {len(product_urls)} products. Starting scrape...")

            # Step 3
            for index, p_url in enumerate(product_urls):
                try:
                    product_data = scrape_product_details(driver, p_url, cat_name)
                    
                    if product_data:
                        all_extracted_data.append(product_data)
                        print(f"[{index+1}/{len(product_urls)}] -> {product_data['Product Name']} -> {product_data['SKU']} -> {p_url}")

                        if len(all_extracted_data) % SAVE_INTERVAL == 0:
                            print(f"   [Autosave] Saving {len(all_extracted_data)} records...")
                            pd.DataFrame(all_extracted_data).to_excel(OUTPUT_FILE, index=False)

                except Exception as e:
                    print(f"   > Error scraping {p_url}: {e}")
                    continue

    except KeyboardInterrupt:
        print("\n\n!!! Script Interrupted by User (Ctrl+C) !!!")
        print("Saving collected data so far...")
    
    except Exception as e:
        print(f"\n!!! Critical Error: {e}")
    
    finally:
        if all_extracted_data:
            df = pd.DataFrame(all_extracted_data)
            df = df.drop_duplicates(subset=["Product URL"])
            df.to_excel(OUTPUT_FILE, index=False)
            print(f"\nCOMPLETED. {len(df)} products saved to {OUTPUT_FILE}")
        else:
            print("\nNo data collected.")
        
        driver.quit()

if __name__ == "__main__":
    main()