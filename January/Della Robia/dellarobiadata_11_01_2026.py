import time
import pandas as pd
import signal
import sys
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager

# --- GLOBAL VARIABLES ---
all_product_details = []
output_file = "dellarobbia_data_v4.xlsx"

def save_data():
    if all_product_details:
        df = pd.DataFrame(all_product_details)
        for i in range(1, 5):
            if f"Image {i}" not in df.columns:
                df[f"Image {i}"] = ""

        # Enforce column order
        base_cols = ["Category", "Product URL", "Product Name", "SKU", "Brand", 
                     "Specifications", "Measurements", "Features", "Catalog PDF"]
        image_cols = ["Image 1", "Image 2", "Image 3", "Image 4"]
        
        final_cols = [c for c in base_cols if c in df.columns] + image_cols
        df = df.reindex(columns=final_cols)
        
        df.to_excel(output_file, index=False)
        print(f"\n[SAVED] Data saved to {output_file} ({len(df)} records)")

def signal_handler(sig, frame):
    print("\n\n[!] Script Interrupted. Saving data...")
    save_data()
    sys.exit(0)

signal.signal(signal.SIGINT, signal_handler)

def get_text_after_header(driver, header_text):
    """
    Revised Logic:
    Only returns text if the UL is the IMMEDIATE sibling.
    If the next tag is another <p> (header) or unrelated div, return N/A.
    """
    try:
        # 1. Find the header paragraph
        # We assume headers are in <p> or <strong> tags usually
        header_el = driver.find_element(By.XPATH, f"//p[contains(text(), '{header_text}')]")
        
        # 2. Get the specific NEXT element (tag) immediately after the header
        next_el = header_el.find_element(By.XPATH, "./following-sibling::*[1]")
        
        # 3. Verify if it is actually a list
        if next_el.tag_name.lower() == 'ul':
            return next_el.text.strip()
        else:
            # If the next element is not a UL (e.g. it is the "Measurements" header), stop.
            return "N/A"
    except:
        return "N/A"

def scrape_dellarobbia():
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    wait = WebDriverWait(driver, 10)
    
    base_url = "https://www.dellarobbia.com"
    product_urls_list = []

    try:
        # --- STAGE 1: GATHER URLS ---
        print("--- STAGE 1: Collecting Product URLs ---")
        driver.get(base_url)
        time.sleep(3)

        try:
            collections_menu = wait.until(EC.element_to_be_clickable((By.XPATH, "//nav//a[contains(text(), 'Collections')]")))
            actions = ActionChains(driver)
            actions.move_to_element(collections_menu).perform()
            time.sleep(2)
            parent_li = collections_menu.find_element(By.XPATH, "./parent::li")
            category_elements = parent_li.find_elements(By.CSS_SELECTOR, ".subnav ul li a")
        except Exception as e:
            print(f"Menu Error: {e}")
            return

        categories = [{"name": c.get_attribute("innerText").strip(), "url": c.get_attribute("href")} for c in category_elements if c.get_attribute("href")]
        print(f"Found {len(categories)} categories.")

        for cat in categories:
            print(f"Scanning Category: {cat['name']}")
            driver.get(cat['url'])
            time.sleep(2)

            last_height = driver.execute_script("return document.body.scrollHeight")
            while True:
                driver.execute_script("window.scrollBy(0, 1000);")
                time.sleep(1)
                new_height = driver.execute_script("return document.body.scrollHeight")
                current_scroll = driver.execute_script("return window.scrollY + window.innerHeight")
                if current_scroll >= new_height:
                    break
            time.sleep(2)

            links = driver.find_elements(By.CSS_SELECTOR, "a.sqs-block-image-link")
            for link in links:
                raw_url = link.get_attribute("href")
                if raw_url:
                    full_url = base_url + raw_url if raw_url.startswith("/") else raw_url
                    if not any(p['url'] == full_url for p in product_urls_list):
                        product_urls_list.append({"category": cat['name'], "url": full_url})
            
            print(f"Total products found: {len(product_urls_list)}")

        # --- STAGE 2: SCRAPE DETAILS ---
        print("\n--- STAGE 2: Scraping Details ---")
        total_products = len(product_urls_list)

        for index, item in enumerate(product_urls_list, 1):
            url = item['url']
            category = item['category']
            
            try:
                driver.get(url)
                wait.until(EC.presence_of_element_located((By.TAG_NAME, "h2")))
                time.sleep(1)

                try:
                    product_name = driver.find_element(By.TAG_NAME, "h2").text.strip()
                except: product_name = "N/A"

                try:
                    sku = driver.find_element(By.XPATH, "//div[contains(@class, 'sqs-block-html')]//p/strong").text.strip()
                except: sku = "N/A"

                # New strict extraction logic
                specs = get_text_after_header(driver, "Specifications")
                measurements = get_text_after_header(driver, "Measurements")
                features = get_text_after_header(driver, "Features")

                # PDF Extraction with scroll
                pdf_url = "N/A"
                try:
                    pdf_element = driver.find_element(By.XPATH, "//a[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'catalog pdf')]")
                    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", pdf_element)
                    time.sleep(1) 
                    pdf_url = pdf_element.get_attribute("href")
                except:
                    try:
                        pdf_element = driver.find_element(By.XPATH, "//div[contains(@class, 'sqs-block-button')]//a[contains(@href, '.pdf')]")
                        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", pdf_element)
                        time.sleep(1)
                        pdf_url = pdf_element.get_attribute("href")
                    except:
                        pass

                # Images (Limit 4)
                driver.execute_script("window.scrollTo(0, 0);") 
                image_elements = driver.find_elements(By.CSS_SELECTOR, ".sqs-block-image img")
                image_urls = []
                for img in image_elements:
                    src = img.get_attribute("data-src") or img.get_attribute("src")
                    if src and src not in image_urls:
                        image_urls.append(src)

                print(f"[{index}/{total_products}] {category} -> {product_name} -> {sku}")

                product_data = {
                    "Category": category,
                    "Product URL": url,
                    "Product Name": product_name,
                    "SKU": sku,
                    "Brand": "Della Robbia",
                    "Specifications": specs,
                    "Measurements": measurements,
                    "Features": features,
                    "Catalog PDF": pdf_url
                }

                for i, img_url in enumerate(image_urls[:4], 1):
                    product_data[f"Image {i}"] = img_url

                all_product_details.append(product_data)

                if index % 50 == 0:
                    save_data()

            except Exception as e:
                print(f"Error scraping {url}: {e}")

    except Exception as global_e:
        print(f"Critical Error: {global_e}")

    finally:
        driver.quit()
        save_data()
        print("\n--- Scraping Completed ---")

if __name__ == "__main__":
    scrape_dellarobbia()