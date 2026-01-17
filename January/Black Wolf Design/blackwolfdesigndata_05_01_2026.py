import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException

# --- HELPER: Save Data ---
def save_data_to_excel(data, filename="BlackWolf_Final_Corrected.xlsx"):
    if not data:
        return
    df = pd.DataFrame(data)
    # Define Column Order
    columns = [
        "Category", "Product Name", "Product URL", "SKU", "Brand", 
        "Description", "Width", "Depth", "Height", "Diameter", 
        "Tearsheet", "Image1", "Image2", "Image3", "Image4"
    ]
    # Ensure all columns exist, fill missing with N/A
    for col in columns:
        if col not in df.columns:
            df[col] = "N/A"
            
    df = df[columns] 
    df.to_excel(filename, index=False)
    print(f"    [Saved] Data saved to {filename} ({len(data)} records)")

# --- STEP 1: Get Categories ---
def get_all_categories(driver):
    category_list = []
    base_url = "https://www.blackwolfdesign.com"

    print("--- STEP 1: Extracting Categories from Menu ---")
    driver.get(base_url)
    wait = WebDriverWait(driver, 15)

    try:
        furniture_menu_id = "jet-mega-menu-item-1286"
        furniture_menu = wait.until(EC.presence_of_element_located((By.ID, furniture_menu_id)))

        actions = ActionChains(driver)
        actions.move_to_element(furniture_menu).perform()
        
        mega_container = wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "jet-mega-menu-mega-container")))
        time.sleep(2) 

        raw_links = mega_container.find_elements(By.TAG_NAME, "a")
        unique_urls = set()

        for link in raw_links:
            try:
                name = link.text.strip()
                url = link.get_attribute("href")

                if url and "product-category" in url:
                    if url.startswith("/"):
                        url = base_url + url
                    clean_url = url.rstrip("/")
                    
                    if clean_url not in unique_urls:
                        unique_urls.add(clean_url)
                        if not name: 
                            name = clean_url.split("/")[-1].replace("-", " ").title()
                        category_list.append({"name": name, "url": url})

            except StaleElementReferenceException:
                continue

    except Exception as e:
        print(f"Error getting categories: {e}")

    print(f" > Total unique categories found: {len(category_list)}\n")
    return category_list

# --- STEP 2: Extract Product Details (Fixed Logic) ---
def get_product_details(driver, product_url):
    details = {
        "Breadcrumb_Raw": "N/A", 
        "Product Name": "N/A", 
        "SKU": "N/A",
        "Brand": "Black Wolf Design", 
        "Description": "N/A",
        "Width": "N/A",
        "Depth": "N/A",
        "Height": "N/A",
        "Diameter": "N/A",
        "Tearsheet": "N/A",
        "Image1": "N/A",
        "Image2": "N/A",
        "Image3": "N/A",
        "Image4": "N/A"
    }
    
    try:
        driver.get(product_url)
        # Wait for page load
        WebDriverWait(driver, 10).until(lambda d: d.find_elements(By.TAG_NAME, "h1") or d.find_elements(By.CLASS_NAME, "jet-woo-product-gallery"))
    except:
        return details 

    # --- 1. FIXED Breadcrumb Extraction (No extra slashes) ---
    try:
        # Find the "Home" link
        home_link = driver.find_element(By.XPATH, "//a[text()='Home']")
        
        # Get the container (3 levels up usually captures the whole breadcrumb row)
        breadcrumb_container = home_link.find_element(By.XPATH, "./ancestor::div[contains(@class, 'elementor-element')][3]")
        
        # Get raw text
        raw_text = breadcrumb_container.text
        
        # CLEANING LOGIC:
        # 1. Split by newline to get parts
        # 2. Filter out empty lines AND lines that are just "/" or "|"
        parts = [
            line.strip() 
            for line in raw_text.split('\n') 
            if line.strip() and line.strip() not in ["/", "\\", "|", ">"]
        ]
        
        # 3. Join with clean " / "
        clean_bread = " / ".join(parts)
        
        if clean_bread.startswith("Home"):
            details["Breadcrumb_Raw"] = clean_bread
    except:
        # Fallback: Build from URL if Elementor fails
        try:
            url_parts = product_url.strip("/").split("/")
            if "product-category" in url_parts:
                idx = url_parts.index("product-category")
                cats = [p.replace("-", " ").title() for p in url_parts[idx+1:]]
                details["Breadcrumb_Raw"] = "Home / Furniture / " + " / ".join(cats)
        except:
            pass

    # --- 2. Product Name and SKU ---
    try:
        title_element = driver.find_element(By.CSS_SELECTOR, "h1.product_title, h1.entry-title")
        raw_text = title_element.text.strip()
        if raw_text:
            details["SKU"] = raw_text
            details["Product Name"] = raw_text
    except:
        pass

    # Fallback Name from URL
    if details["Product Name"] == "N/A":
        try:
            slug = product_url.rstrip("/").split("/")[-1]
            if slug:
                clean_slug = slug.replace("-", " ").title()
                details["Product Name"] = clean_slug
                details["SKU"] = clean_slug
        except:
            pass

    # --- 3. Description ---
    try:
        desc_el = driver.find_element(By.CLASS_NAME, "woocommerce-product-details__short-description")
        details["Description"] = desc_el.text.strip().replace("\n", " ")
    except:
        pass

    # --- 4. Dimensions ---
    try:
        attr_items = driver.find_elements(By.CLASS_NAME, "product-attr--item")
        for item in attr_items:
            try:
                name_el = item.find_element(By.CLASS_NAME, "product-attr--name")
                val_el = item.find_element(By.CLASS_NAME, "product-attr--dimention")
                
                label = name_el.text.strip().lower()
                value = val_el.text.strip()
                
                if "width" in label:
                    details["Width"] = value
                elif "depth" in label:
                    details["Depth"] = value
                elif "height" in label:
                    details["Height"] = value
                elif "dia" in label or "diameter" in label:
                    details["Diameter"] = value
            except:
                continue
    except:
        pass
        
    # --- 5. FIXED Tearsheet Extraction (Broader Search) ---
    try:
        tearsheet_btn = driver.find_element(By.XPATH, "//a[contains(., 'Download Tear Sheet')]")
        ts_url = tearsheet_btn.get_attribute("href")
        if ts_url:
            details["Tearsheet"] = ts_url
    except:
        pass

    # --- 6. Images ---
    unique_images = []
    
    try:
        active_slide = driver.find_element(By.CSS_SELECTOR, ".jet-woo-product-gallery__image-item.swiper-slide-active .jet-woo-product-gallery__image-link")
        img_url = active_slide.get_attribute("href")
        if img_url:
            unique_images.append(img_url)
    except:
        pass

    try:
        grid_links = driver.find_elements(By.CSS_SELECTOR, ".jet-woo-product-gallery__content .jet-woo-product-gallery__image-link")
        for link in grid_links:
            url = link.get_attribute("href")
            if url and url not in unique_images:
                unique_images.append(url)
    except:
        pass

    for i in range(4):
        key = f"Image{i+1}"
        if i < len(unique_images):
            details[key] = unique_images[i]

    return details

# --- MAIN SCRAPER ---
def scrape_blackwolf_targeted():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    # options.add_argument("--headless") 
    driver = webdriver.Chrome(options=options)
    
    final_data = []
    total_counter = 0

    print("--- Starting Scraper ---")
    print("Press Ctrl+C at any time to stop and save progress.")

    try:
        categories = get_all_categories(driver)

        for cat_idx, cat in enumerate(categories):
            cat_name = cat['name']
            cat_url = cat['url']
            
            print(f"\n--- Category {cat_idx + 1}/{len(categories)}: {cat_name} ---")
            driver.get(cat_url)
            time.sleep(3)

            # Load More Logic
            while True:
                try:
                    load_more_wrapper = driver.find_element(By.ID, "cat-load-more")
                    if not load_more_wrapper.is_displayed():
                        break
                    load_more_btn = load_more_wrapper.find_element(By.TAG_NAME, "a")
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", load_more_btn)
                    time.sleep(1)
                    driver.execute_script("arguments[0].click();", load_more_btn)
                    time.sleep(2)
                except:
                    break
            
            # Find Products
            temp_product_urls = []
            try:
                grids = driver.find_elements(By.CLASS_NAME, "jet-listing-grid__items")
                target_grid = max(grids, key=lambda g: len(g.find_elements(By.CLASS_NAME, "jet-listing-grid__item")), default=None)
                
                if target_grid:
                    items = target_grid.find_elements(By.CLASS_NAME, "jet-listing-grid__item")
                    for item in items:
                        try:
                            link_el = item.find_element(By.TAG_NAME, "a")
                            p_url = link_el.get_attribute("href")
                            
                            p_name_log = "Processing..."
                            
                            if p_url and "product" in p_url:
                                temp_product_urls.append({"name": p_name_log, "url": p_url})
                        except:
                            continue
            except Exception as e:
                print(f"Error extracting URLs for category: {e}")

            print(f"    > Found {len(temp_product_urls)} products. Starting Detail Scraping...")

            # Scrape Details
            for p_idx, prod in enumerate(temp_product_urls):
                p_url = prod['url']
                
                total_counter += 1
                print(f"Scraping {p_idx+1}/{len(temp_product_urls)} (Total: {total_counter}) -> {p_url}")

                details = get_product_details(driver, p_url)

                # --- CATEGORY ASSIGNMENT ---
                final_category = details["Breadcrumb_Raw"]
                if final_category == "N/A" or "Privacy Policy" in final_category:
                    final_category = f"Home / Furniture / {cat_name} / {details['Product Name']}"

                row = {
                    "Category": final_category,
                    "Product URL": p_url,
                    **details 
                }
                
                if "Breadcrumb_Raw" in row:
                    del row["Breadcrumb_Raw"]

                final_data.append(row)

                # SAVE every 50 products
                if len(final_data) % 50 == 0:
                    print("    >>> Checkpoint: Saving data...")
                    save_data_to_excel(final_data)

    except KeyboardInterrupt:
        print("\n\n!!! Script Interrupted by User (Ctrl+C) !!!")
        print("Saving collected data before exiting...")
    
    except Exception as e:
        print(f"\nCRITICAL ERROR: {e}")
    
    finally:
        driver.quit()
        if final_data:
            save_data_to_excel(final_data)
            print(f"\nDone. Successfully saved {len(final_data)} products.")
        else:
            print("No data collected.")

if __name__ == "__main__":
    scrape_blackwolf_targeted()
