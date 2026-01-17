import time
import pandas as pd
import sys
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException

# Output Filename
OUTPUT_FILE = "na_furniture_complete_data.xlsx"

def get_all_product_urls(driver):
    """
    Step 1: Navigate menu and collect all product URLs from all categories.
    """
    all_product_links = []
    
    try:
        print("--- STEP 1: Collecting Product URLs ---")
        driver.get("https://www.nafurniture.com/")
        wait = WebDriverWait(driver, 15)

        # 1. Get Categories
        products_menu = wait.until(EC.presence_of_element_located((By.ID, "menu-item-343")))
        actions = ActionChains(driver)
        actions.move_to_element(products_menu).perform()
        time.sleep(2)

        sub_menu_items = products_menu.find_elements(By.CSS_SELECTOR, "ul.sub-menu > li > a")
        categories = []
        for item in sub_menu_items:
            name = item.get_attribute("innerText").strip()
            url = item.get_attribute("href")
            # Exclude Hospitality
            if "HOSPITALITY" in name.upper():
                continue
            if url and name:
                categories.append({"name": name, "url": url})
        
        print(f"Found {len(categories)} categories.")

        # 2. Iterate Categories to get Product Links
        for cat in categories:
            print(f"Scanning Category: {cat['name']}...")
            driver.get(cat['url'])
            
            while True:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(2)

                product_elements = driver.find_elements(By.CSS_SELECTOR, "li.product a.woocommerce-LoopProduct-link")
                
                for link in product_elements:
                    p_url = link.get_attribute("href")
                    if p_url:
                        all_product_links.append({
                            "Category": cat['name'],
                            "Product URL": p_url
                        })

                try:
                    next_button = driver.find_element(By.CSS_SELECTOR, "a.next.page-numbers")
                    next_url = next_button.get_attribute("href")
                    if next_url:
                        driver.get(next_url)
                        time.sleep(2)
                    else:
                        break
                except NoSuchElementException:
                    break
        
        return all_product_links

    except Exception as e:
        print(f"Error gathering links: {e}")
        return []

def scrape_product_details(driver, product_data):
    """
    Step 2: Visit each URL, CLICK ACCORDIONS, and extract detailed info.
    """
    extracted_data = []
    total_products = len(product_data)
    
    print(f"\n--- STEP 2: Extracting Details for {total_products} Products ---")
    print("Press 'Ctrl + C' to stop safely at any time.\n")

    try:
        for index, item in enumerate(product_data):
            current_num = index + 1
            cat_name = item['Category']
            url = item['Product URL']
            
            driver.get(url)
            # Scroll down to ensure accordions are in view (helper for clicking)
            driver.execute_script("window.scrollTo(0, 500);")
            time.sleep(1) 

            # Initialize Row
            row = {
                "Category": cat_name,
                "Product Name": "N/A",
                "Product URL": url,
                "SKU": "N/A",
                "Brand": "Nathan Anthony Furniture",
                "Description": "N/A",
                "Technical Information": "N/A",
                "Dimensions": "N/A",
                "Tearsheet URL": "N/A",
                # Fabric Specific
                "Fabric Name/Color": "N/A",
                "Fabric Grade": "N/A",
                "Fabric Type": "N/A",
                "Fabric Content": "N/A",
                "Fabric Repeat": "N/A",
                "Fabric Width": "N/A",
                "Fabric Direction": "N/A",
                "Fabric Abrasion": "N/A",
                "Fabric Cleaning Code": "N/A",
                # Images
                "Image1": "N/A",
                "Image2": "N/A",
                "Image3": "N/A",
                "Image4": "N/A"
            }

            try:
                # --- 1. Product Name ---
                try:
                    row["Product Name"] = driver.find_element(By.CSS_SELECTOR, "h1.product_title").text.strip()
                except: pass

                print(f"[{current_num}/{total_products}] {cat_name} -> {row['Product Name']} -> {url}")

                # --- 2. Technical Information (CLICK ACCORDION) ---
                try:
                    # Find the toggle using the data-target attribute you provided
                    tech_toggle = driver.find_element(By.CSS_SELECTOR, "h3[data-target='na-product-box-technical']")
                    # Use JS Click (safer than standard click for elements that might be blocked)
                    driver.execute_script("arguments[0].click();", tech_toggle)
                    time.sleep(0.5) # Wait for animation
                    # Now extract the content
                    row["Technical Information"] = driver.find_element(By.ID, "na-product-box-technical").text.strip()
                except: pass

                # --- 3. Dimensions (CLICK ACCORDION) ---
                try:
                    # Find the toggle
                    dim_toggle = driver.find_element(By.CSS_SELECTOR, "h3[data-target='na-product-box-dimensions']")
                    # Click
                    driver.execute_script("arguments[0].click();", dim_toggle)
                    time.sleep(0.5) # Wait for animation
                    # Extract content
                    row["Dimensions"] = driver.find_element(By.ID, "na-product-box-dimensions").text.strip()
                except: pass

                # --- 4. Tearsheet ---
                try:
                    tearsheet_elem = driver.find_element(By.CSS_SELECTOR, "#na-product-box-list-icon-tearsheet a")
                    row["Tearsheet URL"] = tearsheet_elem.get_attribute("href")
                except NoSuchElementException:
                    try:
                        tearsheet_elem = driver.find_element(By.CSS_SELECTOR, "ul.na-product-box-list li a[href*='download=pdf']")
                        row["Tearsheet URL"] = tearsheet_elem.get_attribute("href")
                    except: pass

                # ==========================================================
                # CONDITIONAL LOGIC
                # ==========================================================
                is_fabric = "FABRICS" in cat_name.upper() or "FINISHES" in cat_name.upper()

                if is_fabric:
                    # FABRICS: Parse Description text into columns
                    try:
                        desc_box = driver.find_element(By.CLASS_NAME, "na-product-box-desc")
                        desc_text = desc_box.text
                        
                        lines = desc_text.split('\n')
                        for line in lines:
                            if ":" in line:
                                parts = line.split(":", 1)
                                key = parts[0].strip().upper()
                                val = parts[1].strip()

                                if "NAME/COLOR" in key: row["Fabric Name/Color"] = val
                                elif "GRADE" in key: row["Fabric Grade"] = val
                                elif "TYPE" in key: row["Fabric Type"] = val
                                elif "CONTENT" in key: row["Fabric Content"] = val
                                elif "REPEAT" in key: row["Fabric Repeat"] = val
                                elif "WIDTH" in key: row["Fabric Width"] = val
                                elif "DIRECTION" in key: row["Fabric Direction"] = val
                                elif "ABRASION" in key: row["Fabric Abrasion"] = val
                                elif "CLEANING CODE" in key: row["Fabric Cleaning Code"] = val
                    except: pass
                    
                    try:
                        img_elem = driver.find_element(By.CSS_SELECTOR, ".woocommerce-product-gallery__image img")
                        src = img_elem.get_attribute("data-large_image") or img_elem.get_attribute("src")
                        row["Image1"] = src
                    except: pass

                else:
                    # STANDARD: Save Description & Multiple Images
                    try:
                        row["Description"] = driver.find_element(By.CLASS_NAME, "na-product-box-desc").text.strip()
                    except: pass

                    image_urls = []
                    # Main Image
                    try:
                        main_img = driver.find_element(By.CSS_SELECTOR, ".woocommerce-product-gallery__image img")
                        src = main_img.get_attribute("data-large_image") or main_img.get_attribute("src")
                        if src: image_urls.append(src)
                    except: pass

                    # Thumbnails
                    try:
                        thumbs = driver.find_elements(By.CSS_SELECTOR, "ol.flex-control-nav li img")
                        for thumb in thumbs:
                            t_src = thumb.get_attribute("src")
                            if "100x100" in t_src:
                                t_src = t_src.replace("-100x100", "") 
                            if t_src and t_src not in image_urls:
                                image_urls.append(t_src)
                    except: pass

                    for i in range(min(4, len(image_urls))):
                        row[f"Image{i+1}"] = image_urls[i]

            except Exception as e:
                print(f"Error extracting fields for {url}: {e}")

            extracted_data.append(row)

            # SAVE EVERY 50 ITEMS
            if current_num % 50 == 0:
                print(f"Saving intermediate data ({current_num} products)...")
                save_data(extracted_data)

    except KeyboardInterrupt:
        print("\n\n!!! Script Interrupted by User (Ctrl+C) !!!")
        print("Saving data collected so far...")
    except Exception as e:
        print(f"\nCRITICAL ERROR: {e}")
    finally:
        save_data(extracted_data)

def save_data(data):
    if not data:
        print("No data to save.")
        return
    
    df = pd.DataFrame(data)
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"Data saved to {OUTPUT_FILE}")

def main():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    # options.add_argument("--headless") 
    driver = webdriver.Chrome(options=options)

    try:
        product_links = get_all_product_urls(driver)
        if not product_links:
            print("No products found to scrape.")
            return

        scrape_product_details(driver, product_links)
        
    finally:
        driver.quit()

if __name__ == "__main__":
    main()