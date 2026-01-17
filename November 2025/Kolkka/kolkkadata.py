import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os

# --- CONFIGURATION ---
BASE_URL = "https://www.kolkka.com"
START_URL = "https://www.kolkka.com/collection"
OUTPUT_FILE = "kolkka_complete_data.xlsx"

# --- SELECTORS (Based on your snippets) ---
# Name: h1 with class preFade
SELECTOR_NAME = "//h1[contains(@class, 'preFade')]"
# Size: p tag containing "Size Shown"
SELECTOR_SIZE = "//p[contains(., 'Size Shown')]"
# Full Desc: The container holding the text content
SELECTOR_FULL_DESC = "//div[contains(@class, 'sqs-html-content')]"
# Images: All fluid image containers
SELECTOR_IMAGES = "//div[contains(@class, 'fluid-image-container')]//img"

def init_driver():
    chrome_options = Options()
    # chrome_options.add_argument("--headless") # Uncomment to run invisibly
    driver = webdriver.Chrome(options=chrome_options)
    return driver

def save_data(data_list):
    """Helper to save data to Excel immediately."""
    if not data_list:
        return
    
    df = pd.DataFrame(data_list)
    # Reorder columns for neatness
    cols = [
        'Category', 'Product URL', 'Product Name',  'SKU', 'Brand', 'Size', 
        'Description', 'Full Description HTML', 'Image1', 'Image2', 'Image3', 'Image4'
    ]
    # Ensure all cols exist
    for c in cols:
        if c not in df.columns:
            df[c] = "N/A"
            
    df = df[cols]
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"  [System] Data saved to {OUTPUT_FILE}")

def collect_product_urls(driver):
    """Phase 1: Collect all product URLs first."""
    print("\n--- PHASE 1: Collecting Product URLs ---")
    current_url = START_URL
    visited_urls = set()
    product_urls = []
    page_num = 1
    
    while current_url:
        # Loop Protection
        clean_url = current_url.rstrip('/')
        if clean_url in visited_urls:
            break
        visited_urls.add(clean_url)
        
        print(f"Scanning Page {page_num}...")
        driver.get(current_url)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
        
        # Extract Links
        try:
            links = driver.find_elements(By.XPATH, "//li[contains(@class, 'list-item')]//div[contains(@class, 'list-item-content__description')]//a")
            for link in links:
                href = link.get_attribute('href')
                if href:
                    if href.startswith('/'):
                        href = BASE_URL + href
                    if href not in product_urls:
                        product_urls.append(href)
        except:
            pass
            
        # Next Page Logic
        try:
            next_btn = driver.find_element(By.XPATH, "//a[contains(., 'Next')]")
            next_url = next_btn.get_attribute('href')
            if next_url and next_url.rstrip('/') not in visited_urls:
                current_url = next_url
                page_num += 1
            else:
                current_url = None
        except:
            current_url = None
            
    print(f"--- Collection Complete. Found {len(product_urls)} products. ---\n")
    return product_urls

def extract_product_details(driver, url):
    """Phase 2: Extract details from a single product page."""
    driver.get(url)
    time.sleep(2) # Wait for fade-in animations
    
    item = {
        "Category": "Collections",
        "Product URL": url,
        "SKU": "N/A",
        "Brand": "Kolkka Furniture",
        "Description": "N/A",
        "Product Name": "N/A",
        "Size": "N/A",
        "Full Description HTML": "N/A",
        "Image1": "N/A",
        "Image2": "N/A",
        "Image3": "N/A",
        "Image4": "N/A"
    }
    
    try:
        # 1. NAME
        try:
            name_el = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, SELECTOR_NAME))
            )
            item["Product Name"] = name_el.text.strip()
        except:
            pass

        # 2. SIZE
        try:
            size_el = driver.find_element(By.XPATH, SELECTOR_SIZE)
            item["Size"] = size_el.text.strip()
        except:
            pass

        # 3. FULL DESCRIPTION HTML
        try:
            # We look for the html content div that specifically contains the H1 (Product Name)
            # to avoid grabbing empty or footer blocks
            desc_el = driver.find_element(By.XPATH, "//div[contains(@class, 'sqs-html-content')][.//h1]")
            item["Full Description HTML"] = desc_el.get_attribute('outerHTML').strip()
        except:
            # Fallback to first generic content block found
            try:
                desc_el = driver.find_element(By.XPATH, SELECTOR_FULL_DESC)
                item["Full Description HTML"] = desc_el.get_attribute('outerHTML').strip()
            except:
                pass

        # 4. IMAGES
        try:
            # Find all images in fluid containers
            images = driver.find_elements(By.XPATH, SELECTOR_IMAGES)
            img_urls = []
            
            for img in images:
                # Squarespace uses data-src often for lazy loading
                src = img.get_attribute('data-src')
                if not src:
                    src = img.get_attribute('src')
                
                if src and src not in img_urls:
                    img_urls.append(src)
            
            # Assign to Image1 - Image4
            for i in range(min(len(img_urls), 4)):
                item[f"Image{i+1}"] = img_urls[i]
                
        except:
            pass

    except Exception as e:
        print(f"Error extracting {url}: {e}")
        
    return item

def main():
    driver = init_driver()
    all_data = []
    
    try:
        # Step 1: Get URLs
        urls = collect_product_urls(driver)
        total_products = len(urls)
        
        print("--- PHASE 2: Extracting Details ---")
        
        # Step 2: Iterate and Extract
        for index, url in enumerate(urls):
            current_count = index + 1
            
            # Extract data
            product_data = extract_product_details(driver, url)
            all_data.append(product_data)
            
            # Log to console
            p_name = product_data.get("Product Name", "Unknown")
            print(f"Scraping {current_count}/{total_products} -> {p_name} -> {url}")
            
            # Incremental Save (Every 50 items)
            if current_count % 50 == 0:
                save_data(all_data)
                
    except KeyboardInterrupt:
        print("\n!!! Script Interrupted by User (Ctrl+C) !!!")
        print("Saving collected data so far...")
        
    except Exception as e:
        print(f"\n!!! Script Crashed: {e} !!!")
        print("Saving collected data so far...")
        
    finally:
        # Final Save on exit (Success, Crash, or Interrupt)
        save_data(all_data)
        driver.quit()
        print("\n--- Scraper Finished ---")

if __name__ == "__main__":
    main()