import time
import pandas as pd
import sys
import signal
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# ------------------- Config -------------------
OPERA_BINARY = r"C:\Users\HP\AppData\Local\Programs\Opera\opera.exe"
USER_DATA_DIR = r"C:\OperaProfileVPN"
CHROMEDRIVER_PATH = r"C:\WebDrivers\chromedriver-win64\chromedriver.exe"
CHROMIUM_MAJOR = 140
OUTPUT_FILE = "massoud_complete_data.xlsx"

# Global list to store data in case of crash
collected_full_data = []

# ------------------- Driver Setup -------------------
def setup_opera():
    options = Options()
    options.binary_location = OPERA_BINARY
    options.add_argument("--start-maximized")
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")
    options.add_argument(f"--user-data-dir={USER_DATA_DIR}")
    options.add_argument("--profile-directory=Default")
    options.add_argument("--remote-debugging-port=0")
    options.add_argument("--disable-extensions")
    options.page_load_strategy = 'eager' 

    service = Service(CHROMEDRIVER_PATH)
    driver = uc.Chrome(service=service, options=options, version_main=CHROMIUM_MAJOR)
    return driver

# ------------------- Helper: Save Data -------------------
def save_data_to_excel():
    if collected_full_data:
        print(f"\n>> Saving {len(collected_full_data)} records to '{OUTPUT_FILE}'...")
        df = pd.DataFrame(collected_full_data)
        df.to_excel(OUTPUT_FILE, index=False)
        print(">> Save Complete.")
    else:
        print("\n>> No data to save.")

# ------------------- Helper: Signal Handler (Ctrl+C) -------------------
def signal_handler(sig, frame):
    print("\n\n>> Ctrl+C detected! Stopping and saving data...")
    save_data_to_excel()
    sys.exit(0)

signal.signal(signal.SIGINT, signal_handler)

# ------------------- Core: Scrape Product Details -------------------
def scrape_product_details(driver, product_url, category_context):
    driver.get(product_url)
    time.sleep(2) 
    
    data = {
        "Category": "",
        "Product URL": product_url,
        "Product Name": "",
        "SKU": "",
        "Brand": "Massoud Furniture",
        "Description": ""
    }

    # 1. Breadcrumb Category
    try:
        crumbs = driver.find_elements(By.CSS_SELECTOR, ".pdp-info__breadcrumbs__list a")
        if crumbs:
            data["Category"] = " > ".join([c.get_attribute("innerText").strip().upper() for c in crumbs])
        else:
            data["Category"] = category_context
    except:
        data["Category"] = category_context

    # 2. Name & SKU (Handles both Furniture and Textile formats)
    try:
        heading_div = driver.find_element(By.CSS_SELECTOR, "div.pdp-info__heading")
        h1_elem = heading_div.find_element(By.TAG_NAME, "h1")
        
        try:
            # FORMAT 1: Furniture (Has SKU span)
            sku_span = h1_elem.find_element(By.CSS_SELECTOR, "span.stylenum")
            sku_text = sku_span.get_attribute("innerText").strip()
            full_h1_text = h1_elem.get_attribute("innerText")
            
            data["SKU"] = sku_text
            data["Product Name"] = full_h1_text.replace(sku_text, "").strip()
            
        except NoSuchElementException:
            # FORMAT 2: Fabrics/Leathers (No SKU span)
            data["Product Name"] = h1_elem.get_attribute("innerText").strip()
            data["SKU"] = "N/A"
            
    except Exception as e:
        data["Product Name"] = "N/A"
        data["SKU"] = "N/A"

    # 3. Description
    try:
        desc_div = driver.find_element(By.CSS_SELECTOR, ".pdp-info__description")
        text = desc_div.get_attribute("innerText").strip()
        data["Description"] = text if text else "N/A"
    except:
        data["Description"] = "N/A"

    # ---------------------------------------------------------
    # 4. Handle Accordions (SKIP IF FABRICS/LEATHERS/TRIMS)
    # ---------------------------------------------------------
    # These categories have accordions open by default. Clicking them would close them.
    # Furniture categories usually need clicking.
    
    skip_click_categories = ["Fabrics", "Leathers", "Trims"]
    
    # We check if the current category matches one of the skip types
    if category_context not in skip_click_categories:
        try:
            # Dimensions
            dims_btn = driver.find_elements(By.XPATH, "//span[contains(@class, 'accordion-heading')][contains(text(), 'Dimensions')]")
            if dims_btn:
                driver.execute_script("arguments[0].click();", dims_btn[0])
                time.sleep(0.5)

            # Standard Features
            feats_btn = driver.find_elements(By.XPATH, "//span[contains(@class, 'accordion-heading')][contains(text(), 'Standard Features')]")
            if feats_btn:
                driver.execute_script("arguments[0].click();", feats_btn[0])
                time.sleep(0.5)

            # Details
            details_btn = driver.find_elements(By.XPATH, "//span[contains(@class, 'accordion-heading')][contains(text(), 'Details')]")
            if details_btn:
                driver.execute_script("arguments[0].click();", details_btn[0])
                time.sleep(0.5)
        except:
            pass 
    # else: If it IS in skip_click_categories, we do nothing and proceed to extraction

    # 5. Extract Tabular Data
    fields_map = {
        "Overall": "N/A", "Inside": "N/A", "Arm": "N/A", "Seat": "N/A", 
        "Back Cushion": "N/A", "Seat Cushion": "N/A", "COM": "N/A", "COL": "N/A", "Throw Pillows": "N/A", 
        "Content": "N/A", "Grade": "N/A", "Direction": "N/A", "Sustainability": "N/A", "Width": "N/A", 
        "Repeats": "N/A", "Cleaning Code": "N/A", "Durability": "N/A", "Origin": "N/A"
    }

    try:
        rows = driver.find_elements(By.CSS_SELECTOR, ".pdp-info__tabular_row")
        for row in rows:
            try:
                cols = row.find_elements(By.XPATH, "./div[contains(@class, 'pdp-info__tabular_col')]")
                if len(cols) >= 2:
                    label = cols[0].get_attribute("innerText").strip()
                    val_text = cols[1].get_attribute("innerText").strip().replace("\n", " ")
                    
                    if label in fields_map and val_text:
                        fields_map[label] = val_text
            except:
                continue
    except:
        pass

    data.update(fields_map)

    # 6. Tearsheet
    generic_ts_url = "https://www.massoudfurniture.com/catalogs/#tear-sheets"
    try:
        ts_link = driver.find_element(By.XPATH, "//a[contains(text(), 'Tear Sheet')]")
        href = ts_link.get_attribute("href")
        
        if href and href.strip() != generic_ts_url:
            data["Tearsheet"] = href.strip()
        else:
            data["Tearsheet"] = "N/A"
    except:
        data["Tearsheet"] = "N/A"

    # 7. Images
    try:
        image_urls = []
        
        # Image 1
        try:
            main_img = driver.find_element(By.CSS_SELECTOR, ".pdp-primary__media-slides__item.slick-current img")
            src = main_img.get_attribute("src")
            if src: image_urls.append(src)
        except:
            pass

        # Image 2+
        try:
            thumbnails = driver.find_elements(By.CSS_SELECTOR, ".pdp-primary__thumbnail-stack .slick-slide:not(.slick-cloned) img")
            for thumb in thumbnails:
                src = thumb.get_attribute("src")
                if src and src not in image_urls:
                    image_urls.append(src)
        except:
            pass

        for i in range(4):
            key = f"Image{i+1}"
            if i < len(image_urls):
                data[key] = image_urls[i]
            else:
                data[key] = "N/A"
                
    except Exception as e:
        for i in range(4): data[f"Image{i+1}"] = "N/A"

    return data

# ------------------- Main Scraping Logic -------------------
def collect_data():
    driver = setup_opera()
    wait = WebDriverWait(driver, 10)
    
    products_to_visit = []
    
    try:
        print(">> Opening Website...")
        driver.get("https://www.massoudfurniture.com/")
        time.sleep(5) 

        # --- PHASE 1: Collect Category URLs ---
        print(">> PHASE 1: Collecting Categories...")
        
        menu_map = {
            "Furniture": "furniture-submenu",
            "Custom Choices": "custom-choices-submenu",
            "Textiles & More": "textiles-&-more-submenu"
        }
        textile_allowed_list = ["Fabrics", "Leathers", "Trims"]
        category_links = []

        for menu_name, submenu_id in menu_map.items():
            try:
                menu_btn = wait.until(EC.element_to_be_clickable((By.XPATH, f"//button[span[contains(text(), '{menu_name.split()[0]}')]]")))
                menu_btn.click()
                time.sleep(1)
                
                submenu_items = driver.find_elements(By.XPATH, f"//ul[contains(@id, '{submenu_id.split('-')[0]}')]//a")
                
                for item in submenu_items:
                    link = item.get_attribute("href")
                    title = item.get_attribute("innerText").strip()
                    
                    if not title or "Overview" in title or not link: continue
                    if "Textiles" in menu_name and title not in textile_allowed_list: continue
                    
                    category_links.append((title, link))
            except Exception as e:
                print(f"Error menu {menu_name}: {e}")

        # --- PHASE 2: Collect Product URLs ---
        print(f"\n>> PHASE 2: Scanning Product URLs from {len(category_links)} Categories...\n")
        
        for cat_name, cat_url in category_links:
            driver.get(cat_url)
            time.sleep(2)
            
            while True:
                products = driver.find_elements(By.XPATH, "//article[contains(@class, 'product-results-item')]//a[contains(@class, 'abs-cover')]")
                for p in products:
                    p_url = p.get_attribute("href")
                    if p_url:
                        products_to_visit.append((cat_name, p_url))
                
                try:
                    next_button = driver.find_element(By.XPATH, "//a[i[contains(@class, 'icon-right-open')]]")
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
                    time.sleep(1)
                    next_button.click()
                    time.sleep(3)
                except:
                    break
        
        products_to_visit = list(set(products_to_visit))
        total_products = len(products_to_visit)
        print(f"\n>> FOUND {total_products} PRODUCTS TO SCRAPE.\n")

        # --- PHASE 3: Deep Scraping ---
        print(">> PHASE 3: Scraping Product Details...\n")
        
        for index, (cat_name, prod_url) in enumerate(products_to_visit, start=1):
            try:
                p_data = scrape_product_details(driver, prod_url, cat_name)
                collected_full_data.append(p_data)
                
                p_name = p_data.get("Product Name", "Unknown")
                print(f"Scraping {index}/{total_products} -> {p_name} -> {prod_url}")

                if index % 50 == 0:
                    save_data_to_excel()

            except Exception as e:
                print(f"Error scraping {prod_url}: {e}")
                continue

    except Exception as e:
        print(f"Critical Error: {e}")
        
    finally:
        save_data_to_excel()
        driver.quit()

if __name__ == "__main__":
    collect_data()