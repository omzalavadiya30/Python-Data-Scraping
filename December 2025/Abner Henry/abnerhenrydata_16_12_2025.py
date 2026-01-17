import time
import pandas as pd
import signal
import sys
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# --- GLOBAL VARIABLES ---
all_product_details = []
scraped_urls = set() # Keeps track of URLs we have already visited
output_filename = "AbnerHenry_Full_Product_Data.xlsx"

# --- SAVE FUNCTION ---
def save_data():
    if all_product_details:
        df = pd.DataFrame(all_product_details)
        # We still drop duplicates just in case, but the logic below prevents them mostly
        df = df.drop_duplicates(subset=["Product URL"])
        df.to_excel(output_filename, index=False)
        print(f"\n[SAVED] {len(df)} unique products saved to {output_filename}")
    else:
        print("\n[INFO] No data to save.")

# --- CTRL+C HANDLER ---
def signal_handler(sig, frame):
    print("\n\n[!] Script Interrupted (Ctrl+C). Saving collected data...")
    save_data()
    sys.exit(0)

signal.signal(signal.SIGINT, signal_handler)

def get_abner_henry_data():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    # options.add_argument("--headless") 
    
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 15)
    base_url = "https://abnerhenry.com/"
    
    try:
        # ---------------------------------------------------------
        # STEP 1: COLLECT CATEGORY URLS
        # ---------------------------------------------------------
        print("1. Navigating to home page to fetch categories...")
        driver.get(base_url)
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        
        # Get hidden menu content
        menu_content = wait.until(EC.presence_of_element_located((By.ID, "e-n-menu-content-1711")))
        raw_links = menu_content.find_elements(By.TAG_NAME, "a")
        
        valid_categories = []
        excluded_keywords = [
            "collaborations", "trade", "vuepoint", "login", 
            "register", "about", "press", "request a trade account",
            "customization", "chair program", "hardwoods", "hardware"
        ]

        for link in raw_links:
            url = link.get_attribute("href")
            name = link.get_attribute("textContent").strip().replace("\n", " ")
            
            # Filter empty names and invalid URLs
            if not name or not url or url == "#" or "javascript" in url:
                continue
            
            if any(ex in name.lower() for ex in excluded_keywords) or \
               any(ex in url.lower() for ex in excluded_keywords):
                continue
            
            if (name, url) not in valid_categories:
                valid_categories.append((name, url))

        print(f"   -> Found {len(valid_categories)} valid categories.")

        # ---------------------------------------------------------
        # STEP 2: ITERATE CATEGORIES
        # ---------------------------------------------------------
        for cat_name, cat_url in valid_categories:
            print(f"\n--- Processing Category: {cat_name} ---")
            driver.get(cat_url)
            time.sleep(3)
            
            # --- HANDLE "LOAD MORE" ---
            if "new-furniture" not in cat_url.lower():
                while True:
                    try:
                        load_more_btn = driver.find_element(By.CSS_SELECTOR, "#infinite-handle button")
                        if load_more_btn.is_displayed():
                            driver.execute_script("arguments[0].click();", load_more_btn)
                            time.sleep(4)
                        else:
                            break
                    except:
                        break

            # Collect Product URLs from the current category page
            product_nodes = driver.find_elements(By.CSS_SELECTOR, ".product a.woocommerce-LoopProduct-link")
            category_product_urls = [node.get_attribute("href") for node in product_nodes]
            category_product_urls = list(set(category_product_urls)) # Remove duplicates on page
            
            total_in_cat = len(category_product_urls)
            print(f"   -> Found {total_in_cat} products listings.")

            # ---------------------------------------------------------
            # STEP 3: SCRAPE PRODUCT DETAILS
            # ---------------------------------------------------------
            new_products_count = 0
            
            for index, p_url in enumerate(category_product_urls, 1):
                # --- DUPLICATE CHECKER ---
                if p_url in scraped_urls:
                    # We already extracted this product in a previous category
                    continue
                
                # Mark as visited immediately
                scraped_urls.add(p_url)
                new_products_count += 1

                try:
                    driver.get(p_url)
                    time.sleep(2) 

                    # --- DATA EXTRACTION ---
                    
                    # 1. Product Name
                    try:
                        name_el = driver.find_element(By.CSS_SELECTOR, "h1.product_title")
                        p_name = name_el.text.strip()
                    except:
                        p_name = "N/A"

                    # 2. Breadcrumb
                    try:
                        breadcrumb_el = driver.find_element(By.CSS_SELECTOR, "nav.woocommerce-breadcrumb")
                        raw_crumb = breadcrumb_el.text.replace("\xa0", " ").strip().upper()
                        category_path = raw_crumb.replace("/", " / ")
                        category_path = " ".join(category_path.split())
                    except:
                        category_path = "N/A"

                    # 3. SKU
                    sku = "N/A"
                    try:
                        paragraphs = driver.find_elements(By.CSS_SELECTOR, "div.woocommerce-product-details__short-description p")
                        for p in paragraphs:
                            text = p.text.strip()
                            clean_text = " ".join(text.split())
                            if not clean_text: continue

                            if len(clean_text) < 15 and re.match(r'^[A-Z]{1,3}\s?[\d-]+[A-Z0-9]*$', clean_text):
                                sku = clean_text
                                break
                            
                            match = re.search(r'\b[A-Z]{1,2}\d{3,5}(?:-[A-Z0-9]+)?\b', clean_text)
                            if match:
                                sku = match.group(0)
                                break
                    except:
                        pass

                    # 4. Brand
                    brand = "Abner Henry"

                    # 5. Description & Dimensions
                    description = "N/A"
                    dimensions = "N/A"
                    try:
                        desc_div = driver.find_element(By.CSS_SELECTOR, "div.woocommerce-product-details__short-description")
                        full_desc = desc_div.text.strip()
                        lines = full_desc.split('\n')
                        cleaned_lines = []
                        
                        for l in lines:
                            l_clean = l.strip()
                            if not l_clean: continue
                            
                            if "Customize This Piece" in l or "Download the Tearsheet" in l: continue
                            
                            # Reject line if too long (likely description)
                            if len(l_clean) > 60:
                                cleaned_lines.append(l)
                                continue

                            # Reject hardware keywords
                            if any(bad_word in l_clean for bad_word in ["Knob", "Pull", "Handle", "Hardware"]):
                                cleaned_lines.append(l)
                                continue

                            # Regex for Dimensions
                            fraction_num = r'\d+(?:\s*[½⅓⅔¼¾])?'
                            
                            # Matches: 64"W x 20"D... OR 30"DIA x 30"H
                            is_dimension = re.search(f'{fraction_num}\\s*(?:[”"″]|in)?\\s*(?:[WwDdHh]|Dia|DIA|Diam)?\\.?'
                                                     f'\\s*[xX]\\s*{fraction_num}', l, re.IGNORECASE) or \
                                           re.search(f'{fraction_num}\\s*(?:[”"″]|in)?\\s*Dia', l, re.IGNORECASE)
                                           
                            if is_dimension:
                                dimensions = l_clean
                                continue 
                            
                            if sku != "N/A" and sku in l: continue
                            cleaned_lines.append(l)
                        
                        description = "\n".join(cleaned_lines).strip()
                    except:
                        pass

                    # 7. Tearsheet URL
                    try:
                        tearsheet_el = driver.find_element(By.CSS_SELECTOR, "a.pdfDownload")
                        tearsheet_url = tearsheet_el.get_attribute("href")
                    except:
                        tearsheet_url = "N/A"

                    # 8. Images
                    images = []
                    
                    # A. Main Image (Method 1)
                    try:
                        main_img_div = driver.find_element(By.CSS_SELECTOR, "div.woocommerce-product-details__images div[data-thumb]")
                        img_anchor = main_img_div.find_element(By.TAG_NAME, "a")
                        main_img_url = img_anchor.get_attribute("href")
                        if main_img_url: images.append(main_img_url)
                    except:
                        try:
                            img_anchor = driver.find_element(By.CSS_SELECTOR, "div.woocommerce-product-gallery__wrapper .woocommerce-product-gallery__image a")
                            main_img_url = img_anchor.get_attribute("href")
                            if main_img_url: images.append(main_img_url)
                        except: pass

                    # B. Carousel Images
                    try:
                        carousel_wrapper = driver.find_element(By.CLASS_NAME, "elementor-image-carousel-wrapper")
                        driver.execute_script("arguments[0].scrollIntoView();", carousel_wrapper)
                        time.sleep(1)
                        slides = carousel_wrapper.find_elements(By.CSS_SELECTOR, "div.swiper-slide:not(.swiper-slide-duplicate) a")
                        for slide in slides:
                            img_url = slide.get_attribute("href")
                            if img_url and img_url not in images:
                                images.append(img_url)
                    except: pass

                    img1 = images[0] if len(images) > 0 else "N/A"
                    img2 = images[1] if len(images) > 1 else "N/A"
                    img3 = images[2] if len(images) > 2 else "N/A"
                    img4 = images[3] if len(images) > 3 else "N/A"

                    # --- CONSOLE LOG ---
                    print(f"[{index}/{total_in_cat}] [NEW] {p_name} -> SKU: {sku} -> Dim: {dimensions} -> {p_url}")

                    # --- STORE DATA ---
                    product_data = {
                        "Category Path": category_path,
                        "Product URL": p_url,
                        "Product Name": p_name,
                        "SKU": sku,
                        "Brand": brand,
                        "Dimensions": dimensions,
                        "Description": description,
                        "Tearsheet": tearsheet_url,
                        "Image1": img1,
                        "Image2": img2,
                        "Image3": img3,
                        "Image4": img4
                    }
                    
                    all_product_details.append(product_data)
                    
                    if len(all_product_details) % 50 == 0:
                        save_data()

                except Exception as e:
                    print(f"Error scraping {p_url}: {e}")
                    all_product_details.append({"Product URL": p_url, "Product Name": "ERROR", "SKU": "N/A"})
                    continue
            
            if new_products_count == 0:
                print("   -> All products in this category were already scraped. Skipping.")

    except Exception as e:
        print(f"Critical Error: {e}")
    
    finally:
        driver.quit()
        save_data()

if __name__ == "__main__":
    get_abner_henry_data()