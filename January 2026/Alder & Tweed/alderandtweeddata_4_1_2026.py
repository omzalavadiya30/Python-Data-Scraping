import pandas as pd
import time
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException

# ==========================================
# CONFIGURATION
# ==========================================
BASE_URL = "https://alderandtweedfurniture.com/"
OUTPUT_FILE = "alder_tweed_products_detailed.xlsx"
SAVE_INTERVAL = 50  # Save data every 50 products

def setup_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    # options.add_argument("--headless") # Keep commented out to see errors
    
    # CRITICAL FIX: prevents header overflow errors
    options.add_argument("--disable-dev-shm-usage") 
    options.add_argument("--no-sandbox")
    
    driver = webdriver.Chrome(options=options)
    return driver

def get_text(element, selector):
    try:
        return element.find_element(By.CSS_SELECTOR, selector).text.strip()
    except:
        return "N/A"

def collect_categories(driver):
    print("--- Collecting Categories ---")
    driver.get(BASE_URL)
    wait = WebDriverWait(driver, 15)
    
    try:
        nav_menu = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "nav[aria-label='Menu'] ul.elementor-nav-menu")))
    except TimeoutException:
        print("Could not find navigation menu.")
        return []

    top_menu_items = nav_menu.find_elements(By.XPATH, "./li")
    category_data = []

    for item in top_menu_items:
        try:
            link_element = item.find_element(By.TAG_NAME, "a")
            menu_text = link_element.text.strip()
            
            if menu_text in ["Login", "Registration", "", "Home", "About Us"]:
                continue

            classes = item.get_attribute("class")
            
            if "menu-item-has-children" in classes:
                driver.execute_script("arguments[0].click();", link_element)
                time.sleep(1) 
                sub_items = item.find_elements(By.CSS_SELECTOR, "ul.sub-menu li a")
                for sub in sub_items:
                    sub_name = sub.text.strip()
                    sub_url = sub.get_attribute("href")
                    if sub_url and sub_name:
                        full_cat_name = f"{menu_text} - {sub_name}"
                        category_data.append({"name": full_cat_name, "url": sub_url})
            else:
                url = link_element.get_attribute("href")
                if url:
                    category_data.append({"name": menu_text, "url": url})
        except Exception as e:
            print(f"Error processing menu item {menu_text}: {e}")

    return category_data

def scrape_product_urls(driver, category_url):
    """Reuse existing logic to get URLs from a category page"""
    driver.get(category_url)
    product_urls = []
    
    layout_type = "unknown"
    try:
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.qodef-woo-product-list")))
        layout_type = "standard"
    except:
        try:
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.elementor-widget-image")))
            layout_type = "shop_by_room"
        except:
            pass

    if layout_type == "standard":
        while True:
            products = driver.find_elements(By.CSS_SELECTOR, "li.product a.woocommerce-LoopProduct-link")
            for p in products:
                u = p.get_attribute("href")
                if u and u not in product_urls: product_urls.append(u)
            
            try:
                next_btn = driver.find_element(By.CSS_SELECTOR, "a.next.page-numbers")
                driver.execute_script("arguments[0].click();", next_btn)
                time.sleep(3)
            except:
                break
    
    elif layout_type == "shop_by_room":
        last_height = driver.execute_script("return document.body.scrollHeight")
        while True:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height: break
            last_height = new_height
        
        images = driver.find_elements(By.CSS_SELECTOR, "div.elementor-widget-image a")
        for img in images:
            u = img.get_attribute("href")
            if u and "/product/" in u and u not in product_urls:
                product_urls.append(u)

    return product_urls

def extract_product_details(driver, url, category_name):
    """Extracts data with retry logic for 'Bad Request' errors"""
    
    # === RETRY LOGIC FOR BAD REQUEST ===
    max_retries = 3
    for attempt in range(max_retries):
        try:
            driver.get(url)
            
            # Check if we hit the "Bad Request" page
            if "Bad Request" in driver.title or "header field exceeds server limit" in driver.page_source:
                # print(f"   [!] Bad Request detected. Clearing cookies & Retrying... (Attempt {attempt+1})")
                driver.delete_all_cookies()
                time.sleep(2)
                continue # Try getting the URL again
            
            # If we are here, the page loaded correctly
            break 
        except Exception as e:
            # print(f"   [!] Connection error: {e}. Retrying...")
            time.sleep(2)

    wait = WebDriverWait(driver, 5)
    data = {
        "Category": category_name,
        "Product URL": url,
        "Brand": "Alder and Tweed",
        "Product Name": "N/A",
        "SKU": "N/A",
        "Fabric Cover": "N/A",
        "Leather Grade": "N/A",
        "Wood Finish": "N/A",
        "Wood Material": "N/A",
        "Metal Finish": "N/A",
        "Marble": "N/A",
        "Product Dimensions": "N/A",
        "Product Weight": "N/A",
        "Image1": "N/A",
        "Image2": "N/A",
        "Image3": "N/A",
        "Image4": "N/A"
    }

    try:
        # 1. Product Name
        try:
            data["Product Name"] = driver.find_element(By.CSS_SELECTOR, "div.elementor-widget-container h2.elementor-heading-title").text.strip()
        except:
            pass
        
        # 2. SKU
        try:
            data["SKU"] = driver.find_element(By.CSS_SELECTOR, "div.woolentor_product_sku_info span.sku").text.strip()
        except:
            pass

        # 3. Product Details Accordion
        try:
            # Locate accordion
            accordion_titles = driver.find_elements(By.CSS_SELECTOR, ".elementor-tab-title")
            target_accordion = None
            for acc in accordion_titles:
                if "Product Details" in acc.text:
                    target_accordion = acc
                    break
            
            if target_accordion:
                # Open if closed
                if "elementor-active" not in target_accordion.get_attribute("class"):
                    driver.execute_script("arguments[0].click();", target_accordion)
                    time.sleep(1)

                # Scrape Table
                spec_rows = driver.find_elements(By.CSS_SELECTOR, ".elementor-tab-content.elementor-active table.eael-data-table tr")
                for row in spec_rows:
                    row_text = row.text.strip()
                    if ":" in row_text:
                        split_data = row_text.split(":", 1)
                        key = split_data[0].strip()
                        val = split_data[1].strip()
                        
                        if "Fabric Cover" in key: data["Fabric Cover"] = val
                        elif "Leather Grade" in key: data["Leather Grade"] = val
                        elif "Wood Finish" in key: data["Wood Finish"] = val
                        elif "Wood Material" in key: data["Wood Material"] = val
                        elif "Metal Finish" in key: data["Metal Finish"] = val
                        elif "Marble" in key: data["Marble"] = val
                        elif "Dimensions" in key: data["Product Dimensions"] = val
                        elif "Weight" in key: data["Product Weight"] = val
        except:
            pass

        # 4. Images
        try:
            main_img = driver.find_element(By.CSS_SELECTOR, ".wl-single-slider.slick-current img")
            data["Image1"] = main_img.get_attribute("src")
        except:
            pass

        try:
            thumbnails = driver.find_elements(By.CSS_SELECTOR, ".woolentor-thumbnails .woolentor-thumb-single img")
            for i, thumb in enumerate(thumbnails[:3]):
                data[f"Image{i+2}"] = thumb.get_attribute("src")
        except:
            pass

    except Exception as e:
        print(f"Error parsing {url}: {e}")

    return data

def save_data(data_list, filename):
    if not data_list: return
    df = pd.DataFrame(data_list)
    df.to_excel(filename, index=False)
    print(f"   [SAVED] {len(data_list)} products saved to {filename}")

def main():
    driver = setup_driver()
    final_data = []
    
    try:
        # Step 1: Get Categories
        categories = collect_categories(driver)
        print(f"\nTotal Categories: {len(categories)}")

        for cat in categories:
            cat_name = cat['name']
            cat_url = cat['url']
            
            # Clear cookies between categories
            driver.delete_all_cookies()
            time.sleep(1)

            print(f"Collecting URLs for: {cat_name}...")
            product_urls = scrape_product_urls(driver, cat_url)
            print(f"Found {len(product_urls)} products. Starting extraction...")
            
            total_products = len(product_urls)
            
            for index, prod_url in enumerate(product_urls):
                try:
                    # PROACTIVE COOKIE CLEARING
                    # Every 25 products, clear cookies to prevent "Bad Request" entirely
                    if index > 0 and index % 25 == 0:
                        driver.delete_all_cookies()
                        time.sleep(1)

                    p_data = extract_product_details(driver, prod_url, cat_name)
                    final_data.append(p_data)
                    
                    print(f"Scrapping {index+1}/{total_products} -> {p_data['Product Name']} -> {p_data['SKU']} -> {prod_url}")
                    
                    if len(final_data) % SAVE_INTERVAL == 0:
                        save_data(final_data, OUTPUT_FILE)

                except KeyboardInterrupt:
                    print("\n[!] Script stopped by user. Saving captured data...")
                    save_data(final_data, OUTPUT_FILE)
                    sys.exit()
                except Exception as e:
                    print(f"Error on product loop: {e}")

    except KeyboardInterrupt:
        print("\n[!] Script stopped by user. Saving captured data...")
    except Exception as e:
        print(f"\n[!] Critical Error: {e}")
    finally:
        save_data(final_data, OUTPUT_FILE)
        driver.quit()
        print("\nScript Finished.")

if __name__ == "__main__":
    main()