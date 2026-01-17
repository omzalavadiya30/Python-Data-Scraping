import time
import pandas as pd
import undetected_chromedriver as uc
import signal
import sys
import re
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# --- GLOBAL SETTINGS ---
OUTPUT_FILE = "gatcreek_complete_data.xlsx"
final_data_list = []

# --- SIGNAL HANDLER ---
def signal_handler(sig, frame):
    print("\n\n[STOPPING] Script interrupted by user!")
    save_data()
    sys.exit(0)

signal.signal(signal.SIGINT, signal_handler)

def save_data():
    """Saves the collected data to Excel."""
    if final_data_list:
        df = pd.DataFrame(final_data_list)
        # UPDATED COLUMNS: Changed "Dimensions (HTML)" to "Dimensions" (Plain Text)
        cols = ["Category", "Product URL", "Product Name", "SKU", "Price", "Brand",  
                "Short Description", "Overview Text", "Full Description (HTML)", "Dimensions"]
        
        img_cols = [c for c in df.columns if "Image" in c]
        existing_cols = [c for c in cols + img_cols if c in df.columns]
        
        df = df[existing_cols]
        df.to_excel(OUTPUT_FILE, index=False)
        print(f"\n[SAVED] {len(final_data_list)} records saved to {OUTPUT_FILE}")
    else:
        print("\n[WARNING] No data to save.")

def setup_driver():
    """Initializes Undetected Chrome Driver."""
    options = uc.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    driver = uc.Chrome(options=options, version_main=142) 
    return driver

# --- HELPER FUNCTIONS ---

def get_menu_categories(driver, base_url):
    print("Collecting categories from Menu...")
    categories = []
    try:
        nav_bar = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "mega-menu")))
        links = nav_bar.find_elements(By.TAG_NAME, "a")
        unique_urls = set()
        
        for link in links:
            try:
                url = link.get_attribute("href")
                name = link.text.strip() or driver.execute_script("return arguments[0].textContent;", link).strip()
                if url and base_url in url and url not in unique_urls:
                    if "shop-by-room" not in url and "learn" not in url:
                        unique_urls.add(url)
                        categories.append({"Category Name": name, "Category URL": url})
            except: continue
        print(f"Found {len(categories)} unique categories.")
        return categories
    except Exception as e:
        print(f"Error extracting menu: {e}")
        return []

def collect_product_urls(driver, category):
    """Scrapes all product URLs from a single category page."""
    urls = []
    driver.get(category['Category URL'])
    
    while True:
        try:
            time.sleep(2)
            try:
                pager = driver.find_element(By.CSS_SELECTOR, ".toolbar__pager")
                driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", pager)
            except:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1)
        except: pass

        products = driver.find_elements(By.CSS_SELECTOR, "li.product-grid-item a.product-grid-item__link")
        if not products: break

        for p in products:
            u = p.get_attribute("href")
            if u: urls.append(u)

        try:
            next_btns = driver.find_elements(By.CSS_SELECTOR, "a.pager__link--next")
            if next_btns and next_btns[0].get_attribute("href"):
                driver.get(next_btns[0].get_attribute("href"))
            else:
                break
        except: break
    
    return urls

# --- SKU EXTRACTION LOGIC ---

def extract_sku(text):
    """
    Extracts SKU from Dimensions tab text.
    
    Format 1 (Multi): "SKUs (S:82140, D:82141...)" 
       -> Logic: Finds content in (), then grabs the FIRST number found (82140).
       
    Format 2 (Standard): "SKU: 81128" -> Returns 81128.
    
    Format 3 (Parens): "Rocker (26010)" -> Returns 26010.
    """
    if not text: return "N/A"
    
    # Format 1: SKUs (S:82140, D:82141...)
    # Step A: Find the text inside the parentheses after "SKUs"
    match_multi = re.search(r'SKUs?[:\s]*\(([^)]+)\)', text, re.IGNORECASE)
    if match_multi:
        inner_content = match_multi.group(1) # e.g. "S:82140, D:82141"
        # Step B: Find the FIRST number in that string
        first_num_match = re.search(r'(\d+)', inner_content)
        if first_num_match:
            return first_num_match.group(1).strip()
    
    # Format 2: SKU: 12345
    match = re.search(r'SKU[:\s]+(\d+)', text, re.IGNORECASE)
    if match: return match.group(1).strip()

    # Format 3: Standalone 5-digit number in parentheses e.g. "Rocker (26010)"
    match = re.search(r'\(\s*(\d{5})\s*\)', text)
    if match: return match.group(1).strip()
    
    return "N/A"

# --- PRODUCT SCRAPER ---

def scrape_product_details(driver, url, index, total, category_from_menu):
    item = {
        "Category": category_from_menu,
        "Product URL": url,
        "Product Name": "N/A",
        "SKU": "N/A",
        "Price": "N/A",
        "Brand": "Gat Creek",
        "Short Description": "N/A",
        "Overview Text": "N/A",
        "Dimensions": "N/A" # Changed to plain text
    }
    for i in range(1, 5): item[f"Image{i}"] = "N/A"

    try:
        driver.get(url)
        
        # 1. Breadcrumb Category (Robust)
        try:
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".breadcrumbs__list")))
            crumb_items = driver.find_elements(By.CSS_SELECTOR, ".breadcrumbs__list > li")
            path_parts = []
            for li in crumb_items:
                raw_text = driver.execute_script("return arguments[0].textContent;", li)
                clean_text = " ".join(raw_text.split())
                if clean_text: path_parts.append(clean_text)
            if path_parts: item["Category"] = " / ".join(path_parts)
        except: pass

        # 2. Product Name
        try: item["Product Name"] = driver.find_element(By.CSS_SELECTOR, "h1.heading--page").text.strip()
        except: pass

        # 3. Price
        try: 
            item["Price"] = driver.find_element(By.CSS_SELECTOR, "span.price").text.strip()
        except: pass

        # 4. Short Description
        try: item["Short Description"] = driver.find_element(By.CSS_SELECTOR, ".product-view__short-description").text.strip()
        except: pass
        
        # Scroll to Tabs
        try:
            driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", driver.find_element(By.ID, "product-tabs"))
            time.sleep(1)
        except: pass

        # Tab 1: Images
        try:
            driver.execute_script("document.getElementById('tab-title-1').click();")
            time.sleep(0.5)
            imgs = driver.find_elements(By.CSS_SELECTOR, "#tab-1 .tab-gallery__link")
            for i, img in enumerate(imgs[:4]):
                item[f"Image{i+1}"] = img.get_attribute("href")
        except: pass

        # Tab 2: Overview (HTML + Text)
        try:
            driver.execute_script("document.getElementById('tab-title-2').click();")
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, "tab-2")))
            overview_el = driver.find_element(By.ID, "tab-2")
            item["Overview Text"] = driver.execute_script("return arguments[0].innerText;", overview_el).strip()
        except: pass

        # Tab 3: Dimensions & SKU
        try:
            driver.execute_script("document.getElementById('tab-title-3').click();")
            time.sleep(0.5)
            dim_el = driver.find_element(By.ID, "tab-3")
            
            # UPDATED: Extract Plain Text for Dimensions
            item["Dimensions"] = driver.execute_script("return arguments[0].innerText;", dim_el).strip()
            
            # Extract SKU (Using new logic)
            item["SKU"] = extract_sku(dim_el.get_attribute("textContent"))
        except: pass

        print(f"[{index}/{total}] -> {item['Product Name']} -> SKU: {item['SKU']} -> {url}")

    except Exception as e:
        print(f"[{index}/{total}] FAIL -> {url} -> {e}")

    return item

# --- MAIN EXECUTION ---

def main():
    base_url = "https://www.gatcreek.com"
    driver = setup_driver()
    
    try:
        # 1. Open Site & Cloudflare Check
        driver.get(base_url)
        print("---------------------------------------------------------")
        print("ACTION REQUIRED: Verify Cloudflare checkbox manually.")
        print("Once the homepage is loaded, PRESS ENTER here.")
        print("---------------------------------------------------------")
        input("Waiting for user input... (Press Enter to start)")

        # 2. Get Categories
        categories = get_menu_categories(driver, base_url)
        
        # 3. Harvest All Product URLs
        print("\n--- Phase 1: Harvesting Product URLs ---")
        unique_products = {} 
        for cat in categories:
            print(f"Scanning Category: {cat['Category Name']}")
            urls = collect_product_urls(driver, cat)
            for u in urls:
                if u not in unique_products:
                    unique_products[u] = cat['Category Name']
            print(f"  > Found {len(urls)} links. Total unique so far: {len(unique_products)}")
            
        product_list = [{"url": u, "cat": c} for u, c in unique_products.items()]
        total_products = len(product_list)
        print(f"\n--- Phase 1 Complete. Found {total_products} unique products. ---")

        # 4. Scrape Details
        print("\n--- Phase 2: Extracting Product Details ---")
        for i, prod in enumerate(product_list):
            count = i + 1
            data = scrape_product_details(driver, prod['url'], count, total_products, prod['cat'])
            final_data_list.append(data)
            
            if count % 50 == 0:
                save_data()

    except Exception as e:
        print(f"CRITICAL ERROR: {e}")
    
    finally:
        print("Closing driver...")
        driver.quit()
        save_data()
        print("Done.")

if __name__ == "__main__":
    main()