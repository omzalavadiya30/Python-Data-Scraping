import time
import pandas as pd
import re
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager

# --- Configuration ---
EXCLUDED_CATEGORIES = ["Origen", "Lua", "Gem"]
OUTPUT_FILE = "adrianahoyos_final_complete.xlsx"

def get_browser():
    """Sets up the Chrome Webdriver with EAGER loading strategy."""
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    # options.add_argument("--headless") # Uncomment to run in background

    options.page_load_strategy = 'eager' 
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    driver.set_page_load_timeout(60) 
    driver.set_script_timeout(60)
    
    return driver

def clean_text(text):
    """Removes HTML tags and 'Loading...' text."""
    if not text:
        return "N/A"
    clean = re.sub(r'<[^>]+>', '', text)
    clean = " ".join(clean.split())
    if "Loading..." in clean:
        return "N/A"
    return clean

def collect_categories(driver):
    """Collects categories from navbar."""
    categories = {}
    print("--- Collecting Categories ---")
    
    # 1. Quick Ship
    try:
        time.sleep(3) 
        quick_ship_btn = driver.find_element(By.ID, "menu-item-265051")
        link_elem = quick_ship_btn.find_element(By.TAG_NAME, "a")
        link = link_elem.get_attribute("href")
        name = link_elem.get_attribute("textContent").strip()
        if not name: name = "Quick Ship"
        categories[name] = link
        print(f"Found Category: {name}")
    except: pass

    # 2. Products Menu
    try:
        products_menu = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "li.btnMega a"))
        )
        ActionChains(driver).move_to_element(products_menu).perform()
        time.sleep(2)
        
        mega_menu = driver.find_element(By.CSS_SELECTOR, ".mega-menu")
        sub_links = mega_menu.find_elements(By.TAG_NAME, "a")
        
        for link in sub_links:
            url = link.get_attribute("href")
            name = link.get_attribute("textContent").strip()
            
            if not name or not url: continue
            
            if any(ex.lower() in name.lower() for ex in EXCLUDED_CATEGORIES):
                if name not in categories:
                    print(f"Skipping excluded category: {name}")
                continue
            
            if "product" in url and name not in categories:
                categories[name] = url
                print(f"Found Category: {name}")
                
    except Exception as e:
        print(f"Error interacting with menu: {e}")
        
    return categories

def get_product_urls(driver, category_name, category_url):
    """Collects URLs with scroll logic."""
    try:
        driver.get(category_url)
    except TimeoutException:
        print(f"Timeout loading category: {category_name}. Stopping page load...")
        driver.execute_script("window.stop();")
    
    time.sleep(3)
    collected_urls = set()
    last_count = 0
    consecutive_no_change = 0
    
    print(f"\n--- Collecting URLs for: {category_name} ---")
    
    while True:
        try:
            products = driver.find_elements(By.CSS_SELECTOR, "li.product")
            for product in products:
                try:
                    link = product.find_element(By.CSS_SELECTOR, "a.woocommerce-LoopProduct-link").get_attribute("href")
                    collected_urls.add(link)
                except: continue
            
            current_count = len(collected_urls)
            
            if products:
                driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'end'});", products[-1])
            else:
                driver.execute_script("window.scrollBy(0, 500);")
                
            time.sleep(2)
            
            if current_count == last_count:
                consecutive_no_change += 1
                if consecutive_no_change >= 3:
                    break
            else:
                consecutive_no_change = 0
            last_count = current_count
            
        except Exception:
            break
            
    print(f"Total products found in {category_name}: {len(collected_urls)}")
    return list(collected_urls)

def handle_popup(driver):
    """Detects and closes the 'Build Your Own Sectional' popup."""
    try:
        popup = driver.find_elements(By.CLASS_NAME, "swal2-popup")
        if popup:
            close_btn = driver.find_elements(By.CLASS_NAME, "swal2-close")
            if close_btn and close_btn[0].is_displayed():
                driver.execute_script("arguments[0].click();", close_btn[0])
                print("  > Closed 'Build Your Own' popup.")
                time.sleep(1) 
    except Exception as e:
        try:
            driver.execute_script("document.querySelector('.swal2-container').remove();")
            print("  > Forced removed popup via JS.")
        except: pass

def extract_product_details(driver, product_url, index, total, category_name):
    """Extracts details with Fail-Safe Category Logic."""
    
    # 1. Page Load
    attempts = 0
    max_attempts = 3
    page_loaded_ok = False
    
    while attempts < max_attempts:
        try:
            driver.get(product_url)
            page_loaded_ok = True
            break
        except TimeoutException:
            print(f"  > Timeout loading {product_url}. Refreshing...")
            driver.execute_script("window.stop();")
            attempts += 1
        except WebDriverException as e:
            print(f"  > Connection error: {e}")
            time.sleep(2)
            attempts += 1
            
    if not page_loaded_ok:
        print(f"Skipping {product_url}")
        return None

    # Handle Popup
    handle_popup(driver)

    # 2. Wait for Content
    end_time = time.time() + 15
    while time.time() < end_time:
        try:
            if driver.find_elements(By.ID, "sku-id") or \
               driver.find_elements(By.CSS_SELECTOR, "div[role='heading']") or \
               driver.find_elements(By.CLASS_NAME, "product_title"):
                break
        except: pass
        time.sleep(0.5)

    driver.execute_script("window.scrollBy(0, 700);")
    time.sleep(2)

    data = {
        "Category": "N/A", 
        "Product URL": product_url,
        "Product Name": "N/A", 
        "SKU": "N/A",
        "Brand": "Adriana Hoyos", 
        "Description": "N/A",
         "Tearsheet": "N/A",
        "ArmHeight_In": "N/A", "ArmHeight_Cm": "N/A",
        "Depth_In": "N/A", "Depth_Cm": "N/A",
        "Height_In": "N/A", "Height_Cm": "N/A",
        "SeatDepth_In": "N/A", "SeatDepth_Cm": "N/A",
        "SeatHeight_In": "N/A", "SeatHeight_Cm": "N/A",
        "SeatWidth_In": "N/A", "SeatWidth_Cm": "N/A",
        "Width_In": "N/A", "Width_Cm": "N/A",
        "Image1": "N/A", "Image2": "N/A", "Image3": "N/A", "Image4": "N/A"
    }

    try:
        # --- Description ---
        desc_text = "N/A"
        try:
            desc_elem = None
            if len(driver.find_elements(By.CSS_SELECTOR, ".et-dynamic-content-woo--product_description")) > 0:
                desc_elem = driver.find_element(By.CSS_SELECTOR, ".et-dynamic-content-woo--product_description")
            elif len(driver.find_elements(By.ID, "description")) > 0:
                desc_elem = driver.find_element(By.ID, "description")
            
            if desc_elem:
                wait_time = 0
                while wait_time < 5:
                    current_text = desc_elem.get_attribute("textContent").strip()
                    if current_text and "Loading" not in current_text:
                        desc_text = current_text
                        break
                    time.sleep(1)
                    wait_time += 1
        except: pass
        data["Description"] = clean_text(desc_text)

        # --- Tearsheet Extraction ---
        ts_url = "N/A"
        try:
            if len(driver.find_elements(By.ID, "pdfDownload")) > 0:
                ts_url = driver.find_element(By.ID, "pdfDownload").get_attribute("href")
            
            if not ts_url or ts_url == "N/A":
                xpath = "//a[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'tearsheet')]"
                links = driver.find_elements(By.XPATH, xpath)
                if links:
                    ts_url = links[0].get_attribute("href")

            if not ts_url or ts_url == "N/A":
                pdf_links = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf')]")
                if pdf_links:
                    ts_url = pdf_links[0].get_attribute("href")

            data["Tearsheet"] = ts_url
        except: pass

        # --- Product Name (WITH RETRY & textContent) ---
        name_val = "N/A"
        retry_count = 0
        name_selectors = [
            ".et_pb_text_4_tb_body .et_pb_text_inner",
            "div[role='heading'][aria-level='1']",
            "h1.product_title",
            "h1"
        ]
        
        while retry_count < 3: 
            for selector in name_selectors:
                try:
                    elements = driver.find_elements(By.CSS_SELECTOR, selector)
                    if elements:
                        temp_name = elements[0].get_attribute("textContent").strip()
                        if temp_name and "Loading" not in temp_name:
                            name_val = temp_name
                            break
                except: continue
            if name_val != "N/A":
                break
            time.sleep(1)
            retry_count += 1
            
        data["Product Name"] = name_val

        # --- SKU (WITH RETRY LOOP) ---
        sku_val = "N/A"
        retry_count = 0
        while retry_count < 3: 
            try:
                if len(driver.find_elements(By.ID, "sku-id")) > 0:
                    sku_temp = driver.find_element(By.ID, "sku-id").get_attribute("textContent").strip()
                    if sku_temp: 
                        sku_val = sku_temp
                        break
                if sku_val == "N/A" and len(driver.find_elements(By.CSS_SELECTOR, ".sku .et_pb_text_inner")) > 0:
                    sku_temp = driver.find_element(By.CSS_SELECTOR, ".sku .et_pb_text_inner").get_attribute("textContent").strip()
                    if sku_temp:
                        sku_val = sku_temp
                        break
            except: pass
            time.sleep(1)
            retry_count += 1
        
        data["SKU"] = sku_val

        # --- VALIDATION: Prevent Name from being SKU ---
        if data["Product Name"] != "N/A" and data["SKU"] != "N/A":
            if data["Product Name"] == data["SKU"]:
                fallback_selectors = [
                    ".et_pb_text_3_tb_body .et_pb_text_inner",
                    ".et_pb_module[role='heading'] .et_pb_text_inner"
                ]
                for selector in fallback_selectors:
                    try:
                        elements = driver.find_elements(By.CSS_SELECTOR, selector)
                        for el in elements:
                            text = el.get_attribute("textContent").strip()
                            if text and text != data["SKU"] and "Loading" not in text:
                                data["Product Name"] = text
                                break
                    except: continue
        
        # --- Category (WITH RETRY + FALLBACK) ---
        cat_val = "N/A"
        retry_count = 0
        
        while retry_count < 3:
            try:
                # Plan A: WooCommerce Breadcrumb (Direct Text)
                if len(driver.find_elements(By.CLASS_NAME, "woocommerce-breadcrumb")) > 0:
                    bread_elem = driver.find_element(By.CLASS_NAME, "woocommerce-breadcrumb")
                    full_text = bread_elem.get_attribute("textContent").strip()
                    full_text = full_text.replace("›", ">").replace("\n", " ")
                    cat_val = " ".join(full_text.split()).upper()
                    break

                # Plan B: DSM Breadcrumbs
                elif len(driver.find_elements(By.CSS_SELECTOR, ".dsm_breadcrumbs_item")) > 0:
                    crumbs = driver.find_elements(By.CSS_SELECTOR, ".dsm_breadcrumbs_item a span[itemprop='name']")
                    current_crumb = driver.find_elements(By.CSS_SELECTOR, ".dsm_breadcrumbs_crumb_current")
                    crumb_text = [c.get_attribute("textContent").strip().upper() for c in crumbs if c.get_attribute("textContent").strip()]
                    if current_crumb:
                        crumb_text.append(current_crumb[0].get_attribute("textContent").strip().upper())
                    if crumb_text:
                        cat_val = " > ".join(crumb_text)
                        break
            except: pass
            
            time.sleep(1)
            retry_count += 1
            
        # *** FAIL-SAFE FALLBACK ***
        # If extraction failed (still N/A), construct it manually
        if cat_val == "N/A":
            p_name = data["Product Name"] if data["Product Name"] != "N/A" else "ITEM"
            # Format: HOME > PRODUCTS > [CATEGORY_NAME] > [PRODUCT_NAME]
            cat_val = f"HOME > PRODUCTS > {category_name.upper()} > {p_name}"
            print(f"  > Generated fallback category for: {p_name}")

        data["Category"] = cat_val

        # --- Dimensions ---
        dim_map = {"AH": "ArmHeight", "D": "Depth", "H": "Height", "SD": "SeatDepth", "SH": "SeatHeight", "SW": "SeatWidth", "W": "Width"}
        try:
            tabs = driver.find_elements(By.CLASS_NAME, "dsm-tab")
            for tab in tabs:
                if "Dimensions" in tab.get_attribute("textContent"):
                    driver.execute_script("arguments[0].click();", tab)
                    time.sleep(1)

            rows = driver.find_elements(By.CSS_SELECTOR, "#myTable table tbody tr")
            if not rows:
                rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")

            for row in rows:
                cols = row.find_elements(By.TAG_NAME, "td")
                if len(cols) >= 3:
                    k = cols[0].get_attribute("textContent").strip()
                    v_in = cols[1].get_attribute("textContent").strip()
                    v_cm = cols[2].get_attribute("textContent").strip()
                    if k in dim_map:
                        data[f"{dim_map[k]}_In"] = v_in
                        data[f"{dim_map[k]}_Cm"] = v_cm
        except: pass

        # --- Images ---
        try:
            data["Image1"] = driver.find_element(By.CSS_SELECTOR, "#image-intiaro img").get_attribute("src")
        except:
            try:
                data["Image1"] = driver.find_element(By.CSS_SELECTOR, ".woocommerce-product-gallery__image img").get_attribute("src")
            except: pass

        try:
            gallery_links = driver.find_elements(By.CSS_SELECTOR, "#gallery1 .et_pb_gallery_item a")
            if not gallery_links:
                gallery_links = driver.find_elements(By.CSS_SELECTOR, ".woocommerce-product-gallery__wrapper .woocommerce-product-gallery__image a")
            
            for i, link in enumerate(gallery_links[:3]):
                data[f"Image{i+2}"] = link.get_attribute("href")
        except: pass

        print(f"Scraping [{index}/{total}] -> {data['Product Name']} -> {data['SKU']} -> {product_url}")

    except Exception as e:
        print(f"Error scraping data for {product_url}: {e}")

    return data

def main():
    driver = get_browser()
    all_data = []
    base_url = "https://adrianahoyos.com/"
    
    try:
        driver.get(base_url)
        categories = collect_categories(driver)
        
        for cat_name, cat_url in categories.items():
            product_urls = get_product_urls(driver, cat_name, cat_url)
            total_products = len(product_urls)
            
            print(f"--- Scraping Details for {cat_name} ({total_products} items) ---")
            
            for i, p_url in enumerate(product_urls, 1):
                try:
                    p_data = extract_product_details(driver, p_url, i, total_products, cat_name)
                    if p_data:
                        all_data.append(p_data)
                    
                    if len(all_data) % 20 == 0:
                        print(">> Saving backup checkpoint...")
                        pd.DataFrame(all_data).to_excel(OUTPUT_FILE, index=False)
                        
                except KeyboardInterrupt:
                    raise
                except Exception as e:
                    print(f"Skipping product due to critical error: {e}")
            
            pd.DataFrame(all_data).to_excel(OUTPUT_FILE, index=False)
            print(f"Completed category: {cat_name}\n")

    except KeyboardInterrupt:
        print("\n!!! Interrupted by User. Saving data... !!!")
    except Exception as e:
        print(f"Critical Error: {e}")
    finally:
        if all_data:
            df = pd.DataFrame(all_data)
            df.to_excel(OUTPUT_FILE, index=False)
            print(f"Final Data Saved to {OUTPUT_FILE} ({len(all_data)} records)")
        else:
            print("No data collected.")
        driver.quit()

if __name__ == "__main__":
    main()