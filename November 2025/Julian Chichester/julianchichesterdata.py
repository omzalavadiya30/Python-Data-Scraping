import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time
import sys
import signal

# Base URL
BASE_URL = "https://julianchichester.com"

def setup_driver():
    """Sets up the Selenium WebDriver."""
    options = webdriver.ChromeOptions()
    # options.add_argument("--headless") # Keep commented out to see progress
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("start-maximized")
    options.add_argument("--log-level=3") 
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def save_data(data_list, filename):
    """Saves the collected data list to an Excel file."""
    if not data_list:
        print("No data to save.")
        return
        
    print(f"\nSaving {len(data_list)} items to {filename}...")
    try:
        # Core columns in specific order
        core_cols = ['Category', 'Product URL', 'Product Name', 'SKU', 'Brand', 'Description', 'Full Description HTML', 'Image1', 'Image2', 'Image3', 'Image4']
        
        # The specific specs user requested
        spec_keys_from_user = [
            "Dimensions", "Bulb Type", "Upholstery", "COM", "Sizes", "Estimated Lead Time", 
            "Carver", "Single", "Seat Height", "Arm Height", "Large", "Standard", 
            "Medium", "Small", "Top Finish", "Top Size", "Base Finish", "Shape", 
            "Base Size", "Base Height"
        ]
        
        # Dynamic keys (in case there are extra ones found that weren't in the list)
        all_found_keys = set()
        for item in data_list:
            all_found_keys.update(item.keys())
        
        # Calculate any extra columns found that were not in core or user spec list
        extra_cols = sorted(list(all_found_keys - set(core_cols) - set(spec_keys_from_user)))
        
        # Final column order: Core -> User Specs -> Any Extra Specs found
        final_columns = core_cols + spec_keys_from_user + extra_cols
        
        df = pd.DataFrame(data_list)
        
        # Reindex forces all columns to appear, filling missing ones with NaN (or empty)
        df = df.reindex(columns=final_columns) 
        
        df.to_excel(filename, index=False)
        print(f"--- Data successfully saved! ---")
    except Exception as e:
        print(f"Error saving data: {e}")

def get_categories(driver, wait):
    """Clicks the 'Menu' button and scrapes all category URLs."""
    print("Navigating to homepage...")
    driver.get(BASE_URL)
    
    # Close cookie banner if present
    try:
        cookie_btn = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.cky-btn-accept, button.cky-btn-reject, a.cc-btn"))
        )
        driver.execute_script("arguments[0].click();", cookie_btn)
    except: pass

    try:
        menu_button_locator = (By.XPATH, "//a[@class='pre-heads' and contains(text(), 'Menu')]")
        menu_button = wait.until(EC.element_to_be_clickable(menu_button_locator))
        driver.execute_script("arguments[0].click();", menu_button)
        print("Clicked 'Menu' button.")
    except Exception as e:
        print(f"Error clicking Menu button: {e}")
        return []

    try:
        panel = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.panel.open")))
        category_elements = panel.find_elements(By.CSS_SELECTOR, "a[href*='/category/']")
        category_urls = sorted(list(set(el.get_attribute('href') for el in category_elements if el.get_attribute('href'))))
        print(f"Found {len(category_urls)} unique category URLs.")
        return category_urls
    except Exception as e:
        print(f"Error finding category links: {e}")
        return []

def get_product_urls_and_categories(driver, wait, category_urls):
    """Visits each category page and gets the category name and all product URLs."""
    products_to_scrape = []
    print("\n--- Collecting all Product URLs from Categories ---")
    for cat_url in category_urls:
        try:
            driver.get(cat_url)
            try:
                cat_name = wait.until(EC.visibility_of_element_located(
                    (By.CSS_SELECTOR, "div.side.left p.link")
                )).text
            except:
                cat_name = "Unknown Category"
                
            print(f"Finding products in category: {cat_name} ({cat_url})")

            product_elements = driver.find_elements(By.CSS_SELECTOR, "a[href*='/product/']")
            raw_urls = set(el.get_attribute('href') for el in product_elements if el.get_attribute('href'))
            
            for prod_url in raw_urls:
                if "julianchichester.com" in prod_url and "cookie" not in prod_url.lower():
                    products_to_scrape.append((cat_name, prod_url))
                
        except Exception as e:
            print(f"--- No products found or error on page {cat_url}: {e}")
            
    unique_products_to_scrape = list(dict.fromkeys(products_to_scrape))
    print(f"--- Found a total of {len(unique_products_to_scrape)} unique products to scrape. ---")
    return unique_products_to_scrape

def scrape_product_details(driver, wait, category_name, product_url):
    """
    Scrapes all required details from a single product page.
    """
    
    spec_keys_to_find = {
        "Dimensions", "Bulb Type", "Upholstery", "COM", "Sizes", "Estimated Lead Time", 
        "Carver", "Single", "Seat Height", "Arm Height", "Large", "Standard", 
        "Medium", "Small", "Top Finish", "Top Size", "Base Finish", "Shape", 
        "Base Size", "Base Height"
    }
    
    driver.get(product_url)
    
    # 1. Initialize Dictionary with Defaults
    product_info = {
        "Category": category_name,
        "Product URL": product_url,
        "Product Name": "N/A",
        "SKU": "N/A",
        "Brand": "Julian Chichester",
        "Description": "N/A",
        "Full Description HTML": "N/A",
        "Image1": "N/A", "Image2": "N/A", "Image3": "N/A", "Image4": "N/A"
    }
    
    # --- FIX: Set all spec keys to "N/A" initially ---
    # This ensures that if the scraper doesn't find them, they still appear in Excel as "N/A"
    for key in spec_keys_to_find:
        product_info[key] = "N/A"

    try:
        # Handle Cookie Banner
        try:
            cookie_btn = WebDriverWait(driver, 2).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.cky-btn-accept, button.cky-btn-reject"))
            )
            driver.execute_script("arguments[0].click();", cookie_btn)
            time.sleep(0.5)
        except: pass

        # Wait for content
        try:
            wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, "div.loader")))
        except: pass
        
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "h1.title")))

        # --- PART 1: MAIN DATA ---
        try:
            name_el = driver.find_element(By.CSS_SELECTOR, "h1.title")
            product_info["Product Name"] = name_el.get_attribute('aria-label') or name_el.text.strip()
        except: pass
        
        try:
            sku_el = driver.find_element(By.CSS_SELECTOR, "span.sku")
            product_info["SKU"] = (sku_el.get_attribute('aria-label') or sku_el.text).replace('SKU', '').strip()
        except: pass
            
        try:
            product_info["Description"] = driver.find_element(By.CSS_SELECTOR, "div.description p").text
        except: pass

        try:
            product_info["Full Description HTML"] = driver.find_element(By.CSS_SELECTOR, "div.product-details").get_attribute('outerHTML')
        except: pass
            
        images = driver.find_elements(By.CSS_SELECTOR, "div.product-thumbs img")
        for i, img in enumerate(images[:4]): 
            product_info[f"Image{i+1}"] = img.get_attribute('src')


        # --- PART 2: SIDEBAR DATA ---
        try:
            spec_button_locator = (By.CSS_SELECTOR, "button[aria-label='Open specifications sidebar']")
            
            if driver.find_elements(*spec_button_locator):
                spec_button = driver.find_element(*spec_button_locator)
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", spec_button)
                time.sleep(0.5)
                driver.execute_script("arguments[0].click();", spec_button)
                
                # Wait for Overlay
                overlay_locator = (By.CSS_SELECTOR, "div.specifications-overlay.is-active")
                wait.until(EC.visibility_of_element_located(overlay_locator))
                
                # Wait for Options
                options_locator = (By.CSS_SELECTOR, "div.specifications-overlay.is-active .option")
                spec_elements = wait.until(EC.visibility_of_any_elements_located(options_locator))
                
                time.sleep(0.5) 
                
                for spec in spec_elements:
                    try:
                        title_el = spec.find_element(By.CSS_SELECTOR, ".title")
                        key = title_el.get_attribute("textContent").strip()
                        
                        # Only update if key is in our list
                        if key in spec_keys_to_find:
                            val = "N/A"
                            class_attr = spec.get_attribute("class")

                            if "grid" in class_attr:
                                list_items = spec.find_elements(By.CSS_SELECTOR, "ul li .value")
                                val_list = [li.get_attribute("textContent").strip() for li in list_items]
                                val = ", ".join(filter(None, val_list))
                            else:
                                val_el = spec.find_element(By.CSS_SELECTOR, ".value")
                                val = val_el.get_attribute("textContent").strip()

                            # Overwrite "N/A" with actual value if found
                            if val:
                                product_info[key] = val
                                
                    except Exception:
                        continue
            else:
                pass 

        except TimeoutException:
            print(f"      (Sidebar timeout for {product_url})")
        except Exception as e:
            print(f"      (Sidebar error: {e})")
            
    except Exception as e:
        print(f"      (Critical error scraping product: {e})")
        
    return product_info

def main():
    all_product_data = []
    output_file = "julian_chichester_full_data.xlsx"

    def signal_handler(sig, frame):
        print("\n--- Ctrl+C detected! Saving scraped data... ---")
        save_data(all_product_data, output_file)
        sys.exit(0)

    signal.signal(signal.SIGINT, signal_handler)

    driver = setup_driver()
    wait = WebDriverWait(driver, 10)
    
    try:
        categories = get_categories(driver, wait)
        if not categories: return

        products_to_scrape = get_product_urls_and_categories(driver, wait, categories)
        if not products_to_scrape: return

        total_products = len(products_to_scrape)
        print(f"\n--- Starting to scrape {total_products} product details ---")

        for i, (cat_name, prod_url) in enumerate(products_to_scrape):
            product_info = scrape_product_details(driver, wait, cat_name, prod_url)
            all_product_data.append(product_info)
            
            print(f"Scraping {i+1}/{total_products} -> {product_info.get('Product Name')} -> SKU: {product_info.get('SKU')}")
            
            if (i + 1) % 50 == 0:
                save_data(all_product_data, output_file)

        print("\n--- Scraping Complete ---")
        save_data(all_product_data, output_file)

    except Exception as e:
        print(f"\nAn unexpected error occurred: {e}")
        save_data(all_product_data, output_file)
        
    finally:
        driver.quit()
        print("\nBrowser closed.")

if __name__ == "__main__":
    main()