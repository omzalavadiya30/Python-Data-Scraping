import time
import pandas as pd
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException

# --- CONFIGURATION ---
BASE_URL = "https://www.ambellahome.com"
OUTPUT_FILE = "ambella_full_data.xlsx"

def setup_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    options.page_load_strategy = 'eager'

    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(60)
    return driver

# --- STEP 1: CATEGORY COLLECTION ---
def get_categories(driver):
    print("--- Collecting Categories ---")
    try:
        driver.get(BASE_URL)
    except TimeoutException:
        driver.execute_script("window.stop();")

    wait = WebDriverWait(driver, 15)
    actions = ActionChains(driver)
    
    target_menus = ["New", "Living", "Dining", "Bedroom", "Office", "Bath"]
    collected_categories = []

    for menu_name in target_menus:
        try:
            xpath = f"//div[@class='nav-bar']//nav//div/a[contains(text(), '{menu_name}')]"
            menu_item = driver.find_element(By.XPATH, xpath)
            
            if not menu_item.is_displayed(): continue

            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", menu_item)
            time.sleep(0.5)

            if menu_name == "New":
                link = menu_item.get_attribute("href")
                if link:
                    full_link = link if link.startswith("http") else BASE_URL + link
                    collected_categories.append({"name": "New", "url": full_link})
            else:
                actions.move_to_element(menu_item).perform()
                time.sleep(2)
                parent_div = menu_item.find_element(By.XPATH, "./..")
                sub_links = parent_div.find_elements(By.CSS_SELECTOR, ".nav-dropdown a")
                
                for sub in sub_links:
                    try:
                        url = sub.get_attribute("href")
                        name = sub.get_attribute("textContent").strip()
                        if url and name and not url.lower().endswith(".pdf") and "javascript" not in url:
                            if name.lower() == menu_name.lower(): continue
                            full_url = url if url.startswith("http") else BASE_URL + url
                            # Format: Living- Sofas
                            cat_name = f"{menu_name}- {name}"
                            collected_categories.append({"name": cat_name, "url": full_url})
                    except: continue
        except Exception as e:
            print(f"Skipping '{menu_name}': {e}")

    print(f"Total Categories Found: {len(collected_categories)}")
    return collected_categories

# --- STEP 2: PRODUCT URL COLLECTION ---
def get_product_urls_for_category(driver, category):
    cat_name = category['name']
    cat_url = category['url']
    
    print(f"Collecting URLs for: {cat_name}")
    try:
        driver.get(cat_url)
        
        # Scroll logic
        last_height = driver.execute_script("return document.body.scrollHeight")
        for _ in range(5): 
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height: break
            last_height = new_height

        items = driver.find_elements(By.TAG_NAME, "gallery-item")
        if not items:
            items = driver.find_elements(By.CSS_SELECTOR, "shopping-gallery a.descriptions")
        
        urls = []
        for item in items:
            try:
                if item.tag_name == "gallery-item":
                    link_el = item.find_element(By.TAG_NAME, "a")
                else:
                    link_el = item
                
                href = link_el.get_attribute("href")
                if href:
                    if not href.startswith("http"): href = BASE_URL + href
                    urls.append(href)
            except: continue
            
        print(f" -> Found {len(urls)} products in {cat_name}")
        return urls
    except Exception as e:
        print(f"Error collecting URLs for {cat_name}: {e}")
        return []

# --- STEP 3: DETAILS EXTRACTION ---
def scrape_product_details(driver, product_url, category_name):
    # Default Dictionary with N/A
    data = {
        "Category": category_name,
        "Product URL": product_url,
        "Product Name": "N/A",
        "SKU": "N/A",
        "Price": "N/A",
        "Brand": "Ambella Home",
        "Description": "N/A",
        "Exterior_W": "N/A", "Exterior_D": "N/A", "Exterior_H": "N/A",
        "Interior_W": "N/A", "Interior_D": "N/A", "Interior_H": "N/A",
        "Volume_LBS": "N/A", "Volume_CU.FT": "N/A",
        "Seat Height_H": "N/A",
        "Image1": "N/A", "Image2": "N/A", "Image3": "N/A", "Image4": "N/A"
    }

    try:
        driver.get(product_url)
        # Wait for name to ensure page load
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "h3.hide-for-small")))
        except: 
            # If name not found, page might be broken, return partial N/A
            return data

        # 1. Product Name
        try:
            data["Product Name"] = driver.find_element(By.CSS_SELECTOR, "h3.hide-for-small").text.strip()
        except: pass

        # 2. SKU & Price
        try:
            sku_box = driver.find_element(By.CSS_SELECTOR, ".sku-box")
            data["SKU"] = sku_box.find_element(By.TAG_NAME, "h4").text.strip()
            data["Price"] = sku_box.find_element(By.TAG_NAME, "span").text.strip()
        except: pass

        # 3. Description
        try:
            data["Description"] = driver.find_element(By.CSS_SELECTOR, "p.product-description").text.strip()
        except: pass

        # 4. Dimensions (Flex Table Parsing)
        try:
            table_cells = driver.find_elements(By.CSS_SELECTOR, "flex-table.product-table cell-item")
            current_header = ""
            
            for cell in table_cells:
                text = cell.text.strip()
                class_attr = cell.get_attribute("class")
                
                # Check if it's a header (Exterior, Interior, Volume, Seat Height)
                if "cell-header" in class_attr:
                    current_header = text
                elif text and ":" in text:
                    # Parse "W: 21"" -> key="W", val="21""
                    parts = text.split(":")
                    if len(parts) == 2:
                        key = parts[0].strip()
                        val = parts[1].strip()
                        
                        # Construct Column Name: Exterior_W, Volume_LBS, etc.
                        col_name = f"{current_header}_{key}"
                        
                        # Map specifically to our target columns to handle typos or extra spaces
                        if col_name in data:
                            data[col_name] = val
        except: pass

        # 5. Images
        # Image 1 (Main)
        try:
            main_img = driver.find_element(By.CSS_SELECTOR, "magic-zoom a.MagicZoom")
            data["Image1"] = main_img.get_attribute("href")
        except: pass

        # Images 2-4 (Alt Images)
        try:
            alt_imgs = driver.find_elements(By.CSS_SELECTOR, "alt-images alt-image")
            img_count = 2
            for alt in alt_imgs:
                if img_count > 4: break
                src = alt.get_attribute("data-zoom-src")
                if src:
                    data[f"Image{img_count}"] = src
                    img_count += 1
        except: pass

    except Exception as e:
        print(f"Error extracting details for {product_url}: {e}")
    
    return data

# --- MAIN EXECUTION ---
if __name__ == "__main__":
    driver = setup_driver()
    all_data = []
    
    try:
        # 1. Get All Categories
        categories = get_categories(driver)
        
        # 2. Iterate Categories
        for cat in categories:
            # Collect URLs for this category
            product_urls = get_product_urls_for_category(driver, cat)
            total_products = len(product_urls)
            
            print(f"--- Starting Details Scraping for {cat['name']} (Total: {total_products}) ---")
            
            for index, p_url in enumerate(product_urls):
                # 3. Scrape Details
                details = scrape_product_details(driver, p_url, cat['name'])
                all_data.append(details)
                
                # Log Progress
                print(f"[{index+1}/{total_products}] {cat['name']} -> {details['Product Name']} -> {details['SKU']}")
                
                # 4. Save Every 50
                if len(all_data) % 50 == 0:
                    print("...Saving Checkpoint (50 items)...")
                    pd.DataFrame(all_data).to_excel(OUTPUT_FILE, index=False)

    except KeyboardInterrupt:
        print("\nScript Interrupted by User (Ctrl+C). Saving collected data...")
    except Exception as e:
        print(f"\nCritical Error: {e}")
    finally:
        # Final Save
        if all_data:
            print(f"Saving Final Data ({len(all_data)} products) to {OUTPUT_FILE}...")
            pd.DataFrame(all_data).to_excel(OUTPUT_FILE, index=False)
            print("Done.")
        else:
            print("No data collected.")
        
        driver.quit()