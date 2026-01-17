import time
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException, TimeoutException

# --- CONFIGURATION ---
BASE_URL = "https://bellemeadesignature.com"
OUTPUT_FILE = "belle_meade_final_data.xlsx"

def setup_driver():
    options = Options()
    options.add_argument("--start-maximized")
    options.page_load_strategy = 'eager' # Keep 'eager' for speed
    driver = webdriver.Chrome(options=options)
    return driver

def save_data(data, filename):
    """Helper to save list of dicts to Excel."""
    if not data:
        return
    df = pd.DataFrame(data)
    try:
        df.to_excel(filename, index=False)
        print(f"  [SAVED] Data autosaved to {filename} ({len(df)} rows)")
    except Exception as e:
        print(f"  [ERROR] Could not save file: {e}")

# ==========================================
# PHASE 1: URL COLLECTION
# ==========================================

def get_category_urls(driver):
    print("--- Phase 1: Detecting Categories ---")
    driver.get(BASE_URL)
    wait = WebDriverWait(driver, 10)
    categories = {}

    try:
        # Hover over PRODUCTS
        products_menu = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[contains(text(), 'PRODUCTS')]")))
        ActionChains(driver).move_to_element(products_menu).perform()
        time.sleep(1.5)

        links = driver.find_elements(By.XPATH, "//a[contains(text(), 'PRODUCTS')]/following-sibling::ul//a[contains(@class, 'dropdown-item')]")
        for link in links:
            if not link.is_displayed(): continue
            name = link.get_attribute("innerText").strip()
            url = link.get_attribute("href")
            if name and url and "VIEW ALL" not in name.upper():
                categories[name] = url
        
        try:
            billiards_link = driver.find_element(By.XPATH, "//a[contains(text(), 'BILLIARDS/GAMING')]")
            categories["BILLIARDS/GAMING"] = billiards_link.get_attribute("href")
        except: pass

    except Exception as e:
        print(f"Error getting categories: {e}")
    
    return categories

def collect_product_urls(driver, cat_name, cat_url):
    """Collects ALL product URLs for a category first."""
    print(f"  Collecting URLs for: {cat_name}...")
    driver.get(cat_url)
    urls = set()
    
    # Billiards Logic
    if "BILLIARDS" in cat_name.upper():
        time.sleep(2)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        items = driver.find_elements(By.CSS_SELECTOR, "div.second-product-title h4 a")
        for item in items:
            urls.add(item.get_attribute("href"))
        return list(urls)

    # Standard Logic
    while True:
        try:
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".product-items")))
        except: break

        driver.execute_script("window.scrollTo(0, document.body.scrollHeight * 0.7);")
        time.sleep(1)
        
        items = driver.find_elements(By.CSS_SELECTOR, ".product-items .product-title h4 a")
        for item in items:
            if item.is_displayed():
                urls.add(item.get_attribute("href"))
        
        # Pagination
        try:
            next_btn = None
            selectors = [
                "//ul[contains(@class, 'pagination')]//a[descendant::*[contains(@class, 'icon-arrow-right')]]",
                "//ul[contains(@class, 'pagination')]//a[contains(., 'Next page')]",
                "(//ul[contains(@class, 'pagination')]//li/a)[last()]"
            ]
            for sel in selectors:
                try:
                    el = driver.find_element(By.XPATH, sel)
                    if "Previous" not in el.get_attribute("innerHTML") and "arrow-left" not in el.get_attribute("innerHTML"):
                        next_btn = el
                        break
                except: continue
            
            if next_btn and next_btn.get_attribute("href"):
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_btn)
                driver.execute_script("arguments[0].click();", next_btn)
                time.sleep(1)
            else:
                break
        except: break
        
    return list(urls)

# ==========================================
# PHASE 2: PRODUCT DETAIL EXTRACTION
# ==========================================

def extract_product_details(driver, product_url):
    driver.get(product_url)
    time.sleep(1.5) # Wait for HTML render
    
    data = {
        "Category Path": "N/A",
        "Product URL": product_url,
        "Product Name": "N/A",
        "SKU": "N/A",
        "Brand": "Belle Meade Signature",
        "Description": "N/A",
        "Specs": "N/A",
        "Tearsheet": "N/A",
        "Image1": "N/A",
        "Image2": "N/A",
        "Image3": "N/A",
        "Image4": "N/A"
    }

    try:
        # 1. Category Path
        try:
            breadcrumbs = driver.find_elements(By.CSS_SELECTOR, "ul.breadcrumb-list li a")
            path_parts = [b.text.strip() for b in breadcrumbs if b.text.strip()]
            if path_parts:
                data["Category Path"] = " / ".join(path_parts)
        except: pass

        # 2. Product Name (FIXED)
        try:
            # Targeted selector for the product title div specifically
            title_div = driver.find_element(By.CSS_SELECTOR, "div.pro-title h2")
            full_text = title_div.get_attribute("innerText")
            
            # Split by "Add to Wishlist" or Newline to remove button text
            clean_name = full_text.split("Add to Wishlist")[0].strip()
            clean_name = clean_name.split("\n")[0].strip() # Double check
            
            data["Product Name"] = clean_name
        except: 
            data["Product Name"] = "Name Not Found"

        # 3. Description
        try:
            desc_div = driver.find_element(By.ID, "desktab-1")
            data["Description"] = desc_div.get_attribute("innerText").strip()
        except: pass

        # 4. Specs
        try:
            specs_tab = driver.find_element(By.CSS_SELECTOR, "li[data-tab='desktab-2']")
            driver.execute_script("arguments[0].click();", specs_tab)
            time.sleep(0.5)
            specs_div = driver.find_element(By.ID, "desktab-2")
            data["Specs"] = specs_div.get_attribute("innerText").strip()
        except: pass

        # 5. Tearsheet
        try:
            ts_link = driver.find_element(By.XPATH, "//a[contains(text(), 'TEAR SHEET')]")
            data["Tearsheet"] = ts_link.get_attribute("href")
        except: pass

        # 6. Images (FIXED: Scroll to thumbnails)
        try:
            # Image 1 (Main)
            img1 = driver.find_element(By.CSS_SELECTOR, ".pro-zoom-slider .slick-list .slick-current img")
            data["Image1"] = img1.get_attribute("src")
            
            # Scroll to Thumbnail container to trigger lazy load
            thumb_container = driver.find_element(By.CSS_SELECTOR, ".slider-nav")
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", thumb_container)
            time.sleep(1) # Wait for lazy load

            # Collect Thumbnails
            thumbs = driver.find_elements(By.CSS_SELECTOR, ".slider-nav .slick-slide img")
            img_count = 2
            for thumb in thumbs:
                src = thumb.get_attribute("src")
                # Only add if distinct from main image and count <= 4
                if src and src != data["Image1"] and img_count <= 4:
                    data[f"Image{img_count}"] = src
                    img_count += 1
        except: pass

    except Exception as e:
        print(f"    Error extracting details: {e}")

    return data

# ==========================================
# MAIN EXECUTION
# ==========================================

def main():
    driver = setup_driver()
    all_product_data = []
    
    try:
        # 1. Get Categories
        categories = get_category_urls(driver)
        print(f"Found {len(categories)} categories.")

        # 2. Iterate Categories
        for cat_name, cat_url in categories.items():
            
            # A. Collect URLs
            product_urls = collect_product_urls(driver, cat_name, cat_url)
            total_products = len(product_urls)
            print(f"\n--- Processing {cat_name}: {total_products} products found ---")

            # B. Scrape Details
            for i, url in enumerate(product_urls, 1):
                try:
                    details = extract_product_details(driver, url)
                    
                    p_name = details.get("Product Name", "N/A")
                    print(f"Scraping {i}/{total_products} of {cat_name} -> {p_name} -> {url}")
                    
                    all_product_data.append(details)

                    # Autosave every 50
                    if len(all_product_data) % 50 == 0:
                        save_data(all_product_data, OUTPUT_FILE)

                except Exception as e:
                    print(f"  Error on product loop {i}: {e}")
                    continue

    except KeyboardInterrupt:
        print("\n\n[!] Script stopped by user (Ctrl+C). Saving data...")
    except Exception as e:
        print(f"\n\n[!] Script crashed: {e}. Saving data...")
    finally:
        save_data(all_product_data, OUTPUT_FILE)
        driver.quit()
        print("Scraping finished.")

if __name__ == "__main__":
    main()