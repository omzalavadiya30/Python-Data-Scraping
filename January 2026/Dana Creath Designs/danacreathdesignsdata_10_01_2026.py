import pandas as pd
import time
import sys
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException

# --- CONFIGURATION ---
OUTPUT_FILE = "danacreath_detailed_products.xlsx"

def setup_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    # options.add_argument("--headless") 
    driver = webdriver.Chrome(options=options)
    return driver

# ==========================================
# 1. CATEGORY COLLECTION
# ==========================================
def get_categories_from_dom(driver):
    categories = []
    print("--- Collecting Menu Links (DOM Method) ---")
    
    # 1. Quick Ship
    try:
        quick_ship_link = driver.find_element(By.CSS_SELECTOR, "li.menu-item-342 > a")
        categories.append({
            "Menu Name": "Quick Ship",
            "Menu URL": quick_ship_link.get_attribute("href")
        })
        print(f"Found: Quick Ship")
    except NoSuchElementException:
        print("Quick Ship menu item not found.")

    # 2. Product Menu
    try:
        product_menu_li = driver.find_element(By.CSS_SELECTOR, "li.menu-item-341")
        submenu_items = product_menu_li.find_elements(By.XPATH, "./ul[contains(@class, 'sub-menu')]/li")
        
        for item in submenu_items:
            try:
                link = item.find_element(By.TAG_NAME, "a")
                url = link.get_attribute("href")
                name = link.get_attribute("textContent").strip()
                
                if "Custom Gallery" in name or "Shade Gallery" in name:
                    continue
                
                nested_ul = item.find_elements(By.CSS_SELECTOR, "ul.sub-menu")
                
                if len(nested_ul) > 0:
                    if url:
                        categories.append({"Menu Name": name, "Menu URL": url})
                        print(f"Found Parent: {name}")
                    
                    children_links = item.find_elements(By.XPATH, "./ul/li/a")
                    for child in children_links:
                        child_name = child.get_attribute("textContent").strip()
                        child_url = child.get_attribute("href")
                        if child_url == url: continue 
                        
                        full_name = f"{name} - {child_name}"
                        categories.append({"Menu Name": full_name, "Menu URL": child_url})
                        print(f"  Found Child: {full_name}")
                else:
                    if url:
                        categories.append({"Menu Name": name, "Menu URL": url})
                        print(f"Found: {name}")

            except Exception:
                continue
    except Exception as e:
        print(f"Error finding Product menu: {e}")

    return categories

# ==========================================
# 2. PRODUCT URL COLLECTION
# ==========================================
def get_product_urls(driver, category_url):
    product_urls = []
    driver.get(category_url)
    
    is_quick_ship = "quick-ship" in category_url
    
    while True:
        try:
            current_page_products = []
            
            if is_quick_ship:
                try:
                    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.jet-listing-grid")))
                    links = driver.find_elements(By.CSS_SELECTOR, "div.jet-listing-grid__item .elementor-widget-image a")
                    for link in links:
                        u = link.get_attribute("href")
                        if u and u not in product_urls:
                            current_page_products.append(u)
                            product_urls.append(u)
                except TimeoutException: pass
            else:
                try:
                    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "ul.products")))
                    links = driver.find_elements(By.CSS_SELECTOR, "ul.products li.product a.woocommerce-LoopProduct-link")
                    for link in links:
                        u = link.get_attribute("href")
                        if u and u not in product_urls:
                            current_page_products.append(u)
                            product_urls.append(u)
                except TimeoutException: pass 
            
            if not current_page_products:
                break
            
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1)

            try:
                next_btn = driver.find_element(By.CSS_SELECTOR, "a.next.page-numbers")
                driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", next_btn)
                time.sleep(1)
                next_btn.click()
                time.sleep(3)
            except NoSuchElementException:
                break
                
        except Exception:
            break
            
    return product_urls

# ==========================================
# 3. PRODUCT DETAILS EXTRACTION
# ==========================================
def get_elementor_value_by_label(driver, label_text):
    try:
        xpath = f"//div[contains(@class, 'jet-listing-dynamic-field__content') and contains(., '{label_text}')]/ancestor::div[contains(@class, 'elementor-widget')]/following-sibling::div[contains(@class, 'elementor-widget')]//div[contains(@class, 'jet-listing-dynamic-field__content')]"
        val = driver.find_element(By.XPATH, xpath).text.strip()
        return val
    except NoSuchElementException:
        return "N/A"

def scrape_single_product(driver, product_url):
    driver.get(product_url)
    data = {}
    
    data['Brand'] = "Dana Creath Designs"
    data['Product URL'] = product_url

    # --- Product Name ---
    try:
        name_elem = driver.find_element(By.XPATH, "//div[contains(@class, 'jet-listing-dynamic-field__content') and contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'product description')]")
        raw_text = name_elem.text.strip()
        clean_name = re.sub(r'(?i)^product\s+description\s*[:|-]?\s*', '', raw_text).strip()
        data['Product Name'] = clean_name
        data['Description'] = clean_name
    except:
        data['Product Name'] = "N/A"
        data['Description'] = "N/A"

    # --- Specs ---
    data['Category'] = get_elementor_value_by_label(driver, "Category :")
    data['SKU'] = get_elementor_value_by_label(driver, "Item # :")
    data['Wattage'] = get_elementor_value_by_label(driver, "Wattage :")
    data['Candle'] = get_elementor_value_by_label(driver, "Candle :")
    data['Height'] = get_elementor_value_by_label(driver, "Height :")
    data['Width'] = get_elementor_value_by_label(driver, "Width :")
    data['Depth'] = get_elementor_value_by_label(driver, "Depth :")
    data['Shades'] = get_elementor_value_by_label(driver, "Shades :")
    data['Crate Weight'] = get_elementor_value_by_label(driver, "Crate Weight")
    data['Crate Size'] = get_elementor_value_by_label(driver, "Crate Size :")

    try:
        tearsheet = driver.find_element(By.XPATH, "//a[contains(@href, '.pdf')]")
        data['Tearsheet'] = tearsheet.get_attribute("href")
    except:
        data['Tearsheet'] = "N/A"

    # --- IMAGE EXTRACTION (UPDATED) ---
    image_list = []
    
    # 1. Get Image1 (Main Featured Image)
    try:
        # Based on snippet: <div class="jet-woo-product-gallery__image-item featured ...">
        main_img_elem = driver.find_element(By.CSS_SELECTOR, ".jet-woo-product-gallery__image-item.featured a.jet-woo-product-gallery__image-link")
        main_img_url = main_img_elem.get_attribute("href")
        if main_img_url:
            image_list.append(main_img_url)
    except NoSuchElementException:
        pass

    # 2. Get Image2 - Image4 (Thumbnails)
    try:
        # Scroll to thumbnails to ensure they render
        thumbs_container = driver.find_element(By.CSS_SELECTOR, ".jet-gallery-swiper-thumb")
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", thumbs_container)
        time.sleep(1) # Wait for scroll

        # Extract from data-thumb attribute on the wrapper div
        thumb_items = driver.find_elements(By.CSS_SELECTOR, ".jet-woo-swiper-control-thumbs__item")
        
        for item in thumb_items:
            # According to snippet: <div data-thumb="URL"...>
            url = item.get_attribute("data-thumb")
            
            # Fallback to img src if data-thumb is missing
            if not url:
                try:
                    img = item.find_element(By.TAG_NAME, "img")
                    url = img.get_attribute("data-large_image") or img.get_attribute("src")
                except: pass
            
            # Avoid duplicates (don't add Main Image again if it appears in thumbs)
            if url and url not in image_list:
                image_list.append(url)
                
    except Exception as e:
        # If thumbnails section doesn't exist (single image product), just pass
        pass

    # Assign to Columns
    for i in range(1, 5):
        if len(image_list) >= i:
            data[f'Image{i}'] = image_list[i-1]
        else:
            data[f'Image{i}'] = "N/A"

    return data

# ==========================================
# 4. MAIN CONTROLLER
# ==========================================
def save_data(data_list):
    if not data_list: return
    df = pd.DataFrame(data_list)
    cols = ['Category', 'Product URL', 'Product Name', 'SKU', 'Brand', 'Description', 'Wattage', 'Candle', 
            'Height', 'Width', 'Depth', 'Shades', 'Crate Weight', 'Crate Size', 
            'Tearsheet', 'Image1', 'Image2', 'Image3', 'Image4']
    
    for c in cols:
        if c not in df.columns: df[c] = "N/A"
        
    df = df[cols]
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"\n[System] Data Saved to {OUTPUT_FILE}")

def main():
    driver = setup_driver()
    all_extracted_data = []
    
    try:
        driver.get("https://danacreath.com/")
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "menu-1-650120c")))

        categories = get_categories_from_dom(driver)
        
        for cat in categories:
            cat_menu_name = cat['Menu Name']
            cat_url = cat['Menu URL']
            
            print(f"\n--- Gathering URLs for Menu: {cat_menu_name} ---")
            product_urls = get_product_urls(driver, cat_url)
            total_products = len(product_urls)
            print(f"--> Found {total_products} products.")

            for idx, prod_url in enumerate(product_urls, 1):
                try:
                    details = scrape_single_product(driver, prod_url)
                    
                    if details['Category'] == "N/A":
                        details['Category'] = cat_menu_name
                    
                    all_extracted_data.append(details)
                    
                    print(f"Scraping [{idx}/{total_products}] -> {details.get('Product Name', 'N/A')} -> {details.get('SKU', 'N/A')} -> {prod_url}")

                    if len(all_extracted_data) % 50 == 0:
                        save_data(all_extracted_data)
                        
                except KeyboardInterrupt:
                    raise 
                except Exception as e:
                    print(f"Error scraping {prod_url}: {e}")
                    all_extracted_data.append({
                        "Category": cat_menu_name, 
                        "Product URL": prod_url, 
                        "Product Name": "Error", 
                        "SKU": "N/A"
                    })

    except KeyboardInterrupt:
        print("\n\n[!] Script Interrupted by User (Ctrl+C). Saving data...")
        
    except Exception as e:
        print(f"\n[!] Critical Error: {e}")
        
    finally:
        save_data(all_extracted_data)
        driver.quit()
        print("[System] Script Finished.")

if __name__ == "__main__":
    main()