import time
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# --- GLOBAL SETTINGS ---
BASE_URL = "https://rowefurniture.com"
EXCEL_FILENAME = "rowe_furniture_full_details.xlsx"
BATCH_SIZE = 50  # Save every 50 products

def get_rowe_furniture_data():
    # Setup Chrome Driver
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    # options.add_argument("--headless") # Uncomment for invisible background processing
    
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 15)
    short_wait = WebDriverWait(driver, 3)
    actions = ActionChains(driver)

    all_data = [] # List to store all scraped dictionaries

    # Helper function to save data safely
    def save_data_to_excel():
        if not all_data:
            return
        print(f"\n[SYSTEM] Saving {len(all_data)} collected products to Excel...")
        df = pd.DataFrame(all_data)
        df.to_excel(EXCEL_FILENAME, index=False)
        print("[SYSTEM] Save Complete.\n")

    try:
        # --- PHASE 1: COLLECT CATEGORIES ---
        print(f"Navigating to {BASE_URL}...")
        driver.get(BASE_URL)
        time.sleep(3)

        print("Collecting Categories...")
        products_menu = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[normalize-space()='Products']")))
        actions.move_to_element(products_menu).perform()
        time.sleep(2)
        
        category_elements = driver.find_elements(By.CSS_SELECTOR, "li.subcategory-item a")
        categories_to_scrape = []
        for cat in category_elements:
            name = cat.get_attribute("innerText").strip()
            link = cat.get_attribute("href")
            if link and "http" in link:
                categories_to_scrape.append({"name": name, "url": link})
        
        print(f"Found {len(categories_to_scrape)} categories.")

        # --- PHASE 2: LOOP CATEGORIES ---
        for cat_index, cat in enumerate(categories_to_scrape):
            cat_name = cat['name']
            cat_url = cat['url']
            
            print(f"\n--- Processing Category: {cat_name} ---")
            driver.get(cat_url)
            time.sleep(2)

            # --- PHASE 2a: INFINITE SCROLL (COLLECT URLS) ---
            print("  -> Loading all products in category...")
            while True:
                last_height = driver.execute_script("return document.body.scrollHeight")
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                try:
                    short_wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "infinite-scroll-loader")))
                    WebDriverWait(driver, 20).until(EC.invisibility_of_element_located((By.CLASS_NAME, "infinite-scroll-loader")))
                    time.sleep(1)
                except TimeoutException:
                    new_height = driver.execute_script("return document.body.scrollHeight")
                    if new_height == last_height:
                        break
            
            product_elements = driver.find_elements(By.CSS_SELECTOR, ".product-item .picture a")
            product_urls = list(set([elem.get_attribute("href") for elem in product_elements if elem.get_attribute("href")]))
            
            total_products = len(product_urls)
            print(f"  -> Found {total_products} products. Starting detail extraction...")

            # --- PHASE 3: EXTRACT PRODUCT DETAILS ---
            for i, p_url in enumerate(product_urls):
                try:
                    driver.get(p_url)
                    
                    # 1. Product Name
                    try:
                        p_name = driver.find_element(By.CSS_SELECTOR, ".page-title.product-name h1").text.strip()
                    except:
                        p_name = "N/A"

                    # Log Format
                    print(f"Scrapping {i+1}/{total_products} -> {p_name} -> {p_url}")

                    # 2. Breadcrumb
                    try:
                        bc_items = driver.find_elements(By.CSS_SELECTOR, ".breadcrumb ul li a span, .breadcrumb ul li strong")
                        bc_texts = [x.text.strip() for x in bc_items if x.text.strip()]
                        breadcrumb_str = " / ".join(bc_texts)
                    except:
                        breadcrumb_str = "N/A"

                    # 3. SKU
                    try:
                        sku = driver.find_element(By.CSS_SELECTOR, ".sku .value").text.strip()
                    except:
                        sku = "N/A"

                    # 4. Brand
                    brand = "Rowe Furniture"

                    # 5. Assortment
                    try:
                        assortment = driver.find_element(By.CSS_SELECTOR, ".manufacturers .value").text.strip()
                    except:
                        assortment = "N/A"

                    # 6. Description
                    try:
                        description = driver.find_element(By.CSS_SELECTOR, ".short-description").text.strip()
                    except:
                        description = "N/A"

                    # 7. Images
                    images = []
                    try:
                        main_img = driver.find_element(By.CSS_SELECTOR, "#sevenspikes-cloud-zoom a").get_attribute("href")
                        images.append(main_img)
                    except: pass
                    
                    try:
                        thumbs = driver.find_elements(By.CSS_SELECTOR, ".picture-thumbs-item:not(.slick-cloned) a.thumb-item")
                        for thumb in thumbs:
                            img_link = thumb.get_attribute("data-full-image-url")
                            if img_link and img_link not in images:
                                images.append(img_link)
                    except: pass

                    while len(images) < 4:
                        images.append("N/A")

                    # 8. Specs & Dimensions (HANDLING BOTH FORMATS)
                    specs_map = {}
                    try:
                        # Scroll to specs box to ensure visibility (triggers any lazy loading)
                        spec_box = driver.find_element(By.CSS_SELECTOR, ".product-specs-box")
                        driver.execute_script("arguments[0].scrollIntoView();", spec_box)
                        time.sleep(0.5)

                        rows = driver.find_elements(By.CSS_SELECTOR, ".product-specs-box table tr")
                        for row in rows:
                            cols = row.find_elements(By.TAG_NAME, "td")
                            
                            # LOGIC UPDATE: Check for >= 2 columns to handle both formats
                            if len(cols) >= 2:
                                key = cols[0].text.strip()
                                # Only take the first value (cols[1]), ignore cols[2], cols[3] etc.
                                val = cols[1].text.strip()
                                
                                if key:
                                    specs_map[key] = val
                    except: pass

                    # Construct Data Dictionary with "N/A" defaults
                    product_data = {
                        "Category": breadcrumb_str,
                        "Product URL": p_url,
                        "Product Name": p_name,
                        "SKU": sku,
                        "Brand": brand,
                        "Assortment": assortment,
                        "Description": description,
                        # Map Specs specifically using "N/A" if missing
                        "Weight (LB)": specs_map.get("Weight (LB)", "N/A"),
                        "Seat Height (IN)": specs_map.get("Seat Height (IN)", "N/A"),
                        "Seat Depth (IN)": specs_map.get("Seat Depth (IN)", "N/A"),
                        "Arm Height (IN)": specs_map.get("Arm Height (IN)", "N/A"),
                        "Distance Between Arms (IN)": specs_map.get("Distance Between Arms (IN)", "N/A"),
                        "Depth (IN)": specs_map.get("Depth (IN)", "N/A"),
                        "Height (IN)": specs_map.get("Height (IN)", "N/A"),
                        "Length (IN)": specs_map.get("Length (IN)", "N/A"),
                        "Dimensions (IN)": specs_map.get("Dimensions (IN)", "N/A"),
                        "Shelf Features": specs_map.get("Shelf Features", "N/A"),
                        "Adjustable Floor Glides": specs_map.get("Adjustable Floor Glides", "N/A"),
                        "Kd Construction": specs_map.get("Kd Construction", "N/A"),
                        "Country of Origin": specs_map.get("Country of Origin", "N/A"),
                        "Number Of Back Pillows": specs_map.get("Number Of Back Pillows", "N/A"),
                        "Number Of Cushions": specs_map.get("Number Of Cushions", "N/A"),
                        "Standard Cushion": specs_map.get("Standard Cushion", "N/A"),
                        "Collection": specs_map.get("Collection", "N/A"),
                        # Images
                        "Image1": images[0],
                        "Image2": images[1],
                        "Image3": images[2],
                        "Image4": images[3]
                    }

                    all_data.append(product_data)

                    # Batch Save
                    if len(all_data) % BATCH_SIZE == 0:
                        save_data_to_excel()

                except Exception as e:
                    print(f"Error extracting {p_url}: {e}")

    except KeyboardInterrupt:
        print("\n\n[USER INTERRUPT] Script stopped by user (Ctrl+C).")
        print("Saving collected data before exiting...")
    
    except Exception as e:
        print(f"\n[CRITICAL ERROR] {e}")

    finally:
        save_data_to_excel()
        driver.quit()
        print("Driver closed.")

if __name__ == "__main__":
    get_rowe_furniture_data()