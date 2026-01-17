import pandas as pd
import json
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import sys

# --- 1. Configuration ---
EXCEL_FILENAME = 'interlude_home_product_details.xlsx'
PRODUCTS_PER_AUTOSAVE = 50 

# List of category URLs to scrape
CATEGORIES_TO_SCRAPE = [
    # Cabinets & Chests
    {"name": "Bedside", "url": "https://www.interludehome.com/default/category/cabinets-chests/bedside.html"},
    {"name": "Cabinets", "url": "https://www.interludehome.com/default/category/cabinets-chests/cabinets.html"},
    {"name": "Dressers & Chests", "url": "https://www.interludehome.com/default/category/cabinets-chests/dressers-chests.html"},
    
    # Tables
    {"name": "Bar (Tables)", "url": "https://www.interludehome.com/default/category/tables/bar5908.html"},
    {"name": "Cocktail", "url": "https://www.interludehome.com/default/category/tables/cocktail.html"},
    {"name": "Console", "url": "https://www.interludehome.com/default/category/tables/console.html"},
    {"name": "Desk", "url": "https://www.interludehome.com/default/category/tables/desk.html"},
    {"name": "Dining (Tables)", "url": "https://www.interludehome.com/default/category/tables/dining.html"},
    {"name": "Drink", "url": "https://www.interludehome.com/default/category/tables/drink.html"},
    {"name": "Occasional", "url": "https://www.interludehome.com/default/category/tables/occasional.html"},
    {"name": "Game Tables", "url": "https://www.interludehome.com/default/category/tables/game-tables.html"},
    
    # Seating
    {"name": "Dining Chairs", "url": "https://www.interludehome.com/default/category/seating/dining-chairs.html"},
    {"name": "Counter Stools", "url": "https://www.interludehome.com/default/category/seating/counter-stools.html"},
    {"name": "Bar Stools", "url": "https://www.interludehome.com/default/category/seating/bar-stools.html"},
    {"name": "Occasional chairs", "url": "https://www.interludehome.com/default/category/seating/occasional-chairs.html"},
    {"name": "Benches, Ottomans & Stools (Seating)", "url": "https://www.interludehome.com/default/category/seating/benches-ottomans-stools.html"},
    
    # Quick Ship + Upholstery
    {"name": "Coastal (Upholstery)", "url": "https://www.interludehome.com/default/coastal/upholstery.html"},
    {"name": "Quick Ship", "url": "https://www.interludehome.com/default/category/upholstery2152/quick-ship.html"},
    {"name": "IH Custom Upholstery", "url": "https://www.interludehome.com/default/custom-upholstery.html"},
    {"name": "Chairs (Upholstery)", "url": "https://www.interludehome.com/default/category/upholstery2152/chairs3445.html"},
    {"name": "Beds", "url": "https://www.interludehome.com/default/category/upholstery2152/beds8576.html"},
    {"name": "Sectionals", "url": "https://www.interludehome.com/default/category/upholstery2152/sectionals3389.html"},
    {"name": "Sofas", "url": "https://www.interludehome.com/default/category/upholstery2152/sofas1194.html"},
    {"name": "Benches, Ottomans & Stools (Upholstery)", "url": "https://www.interludehome.com/default/category/upholstery2152/benchesottomansstools.html"},
    
    # Décor
    {"name": "Games", "url": "https://www.interludehome.com/default/category/dcor/games.html"},
    {"name": "Mirrors", "url": "https://www.interludehome.com/default/category/dcor/mirrors5756.html"},
    {"name": "Objets", "url": "https://www.interludehome.com/default/category/dcor/objets3997.html"},
    {"name": "Rugs", "url": "https://www.interludehome.com/default/category/dcor/rugs7357.html"},
    {"name": "Trays", "url":    "https://www.interludehome.com/default/category/dcor/trays8165.html"},
    {"name": "Vessels & Bowls", "url": "https://www.interludehome.com/default/category/dcor/vesselsbowls.html"},
    {"name": "Wall Décor", "url": "https://www.interludehome.com/default/category/dcor/walldcor.html"},
]

# --- 2. Helper Functions ---

def get_driver():
    print("Setting up Chrome driver...")
    service = Service(ChromeDriverManager().install())
    options = webdriver.ChromeOptions()
    options.page_load_strategy = 'eager' # Fast loading
    options.add_argument('--disable-gpu')
    options.add_argument('--window-size=1920,1080')
    options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36')
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def safe_find_text(driver, by, selector, default="N/A"):
    try:
        return driver.find_element(by, selector).text.strip()
    except NoSuchElementException:
        return default

def safe_find_attr(driver, by, selector, attribute, default="N/A"):
    try:
        return driver.find_element(by, selector).get_attribute(attribute).strip()
    except NoSuchElementException:
        return default

def extract_images_separated(driver):
    """
    Returns TWO lists: (thumbnail_urls, main_urls)
    Robustly handles Intiaro 360 sliders by waiting for JS injection.
    """
    thumb_urls = []
    main_urls = []
    
    # --- METHOD 1: Try Intiaro 360 Slider (High Priority for Custom Furniture) ---
    # We check this FIRST or allow it to wait, because these load slower than standard DOM.
    try:
        wait = WebDriverWait(driver, 5) # Wait up to 5 seconds for the slider to render
        
        # Check if the specific Intiaro slider structure exists
        # We look for the slider-image class based on your snippet
        slider_imgs = wait.until(EC.presence_of_all_elements_located(
            (By.CSS_SELECTOR, "ul.slider li.slider-frame img.slider-image")
        ))

        if slider_imgs:
            # The slider usually has 36 frames (0 to 35) for a full rotation.
            # We want to extract 4 distinct angles (Front, Side, Back, Side).
            # 36 frames / 4 images = Step of 9. Indices: 0, 9, 18, 27.
            
            extracted_srcs = []
            
            # Scrape ALL 36 frames first to ensure we have the data
            all_frames_urls = []
            for img in slider_imgs:
                # data-src is usually the reliable one in the snippet you shared
                url = img.get_attribute('data-src')
                if not url:
                    url = img.get_attribute('src')
                
                if url and 'http' in url:
                    all_frames_urls.append(url)
            
            # Now pick the specific angles if we have enough frames
            if len(all_frames_urls) >= 36:
                indices_to_pick = [0, 9, 18, 27] # 0°, 90°, 180°, 270°
                for idx in indices_to_pick:
                    extracted_srcs.append(all_frames_urls[idx])
            elif len(all_frames_urls) > 0:
                # Fallback if less than 36 frames, just take the first few
                extracted_srcs = all_frames_urls[:4]

            if extracted_srcs:
                return extracted_srcs, extracted_srcs
                
    except (TimeoutException, NoSuchElementException):
        # If no 360 slider is found within 5 seconds, proceed to standard logic
        pass

    # --- METHOD 2: Standard Magento JSON (for non-360 products) ---
    try:
        scripts = driver.find_elements(By.CSS_SELECTOR, 'script[type="text/x-magento-init"]')
        for script in scripts:
            content = script.get_attribute('innerHTML')
            if 'mage/gallery/gallery' in content:
                data = json.loads(content)
                for key in data:
                    if 'mage/gallery/gallery' in data[key]:
                        gallery_data = data[key]['mage/gallery/gallery'].get('data', [])
                        for item in gallery_data:
                            t_url = item.get('thumb') or item.get('img')
                            if t_url: thumb_urls.append(t_url)
                            m_url = item.get('full') or item.get('img')
                            if m_url: main_urls.append(m_url)
                
                if thumb_urls or main_urls:
                    return thumb_urls, main_urls
    except Exception:
        pass

    # --- METHOD 3: DOM Fallback (Fotorama) ---
    if not thumb_urls:
        try:
            # Try to find standard fotorama images
            images = driver.find_elements(By.CSS_SELECTOR, '.gallery-placeholder img.fotorama__img')
            for img in images:
                src = img.get_attribute('src')
                if src: 
                    thumb_urls.append(src)
                    main_urls.append(src)
        except: 
            pass

    return thumb_urls, main_urls

def save_to_excel(data_list, filename, is_autosave=False):
    if not data_list: return
    df = pd.DataFrame(data_list)
    if 'Product URL' in df.columns:
        df = df.drop_duplicates(subset=['Product URL']) 
    if is_autosave:
        print(f"\n--- AUTOSAVING {len(df)} products ---")
    else:
        print(f"\n--- FINAL SAVE: {len(df)} products ---")
    df.to_excel(filename, index=False)

# --- 3. Scraping Functions ---

def get_product_urls_from_category(driver, category_name, category_url):
    print(f"\n--- Finding Products in Category: {category_name} ---")
    product_urls = set()
    current_url = category_url
    wait = WebDriverWait(driver, 10)

    try:
        driver.get(current_url)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.breadcrumbs')))
        breadcrumbs = safe_find_text(driver, By.CSS_SELECTOR, 'div.breadcrumbs')
    except: breadcrumbs = "N/A"

    while True:
        if driver.current_url != current_url:
             try: driver.get(current_url)
             except: break
        
        try:
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'a.product-item-link')))
            product_elements = driver.find_elements(By.CSS_SELECTOR, 'a.product-item-link')
            
            if not product_elements: break
            
            for prod in product_elements:
                url = prod.get_attribute('href')
                if url: product_urls.add(url)
            
            try:
                next_button = driver.find_element(By.CSS_SELECTOR, 'a.action.next')
                current_url = next_button.get_attribute('href')
            except NoSuchElementException:
                break 
        except TimeoutException:
            break

    print(f"  Total products found: {len(product_urls)}")
    return list(product_urls), breadcrumbs

def scrape_product_details(driver, category_breadcrumb_str, product_urls, all_data_list, filename):
    total_products = len(product_urls)
    print(f"\n--- Scraping {total_products} Product Details ---")
    wait = WebDriverWait(driver, 5)

    for i, url in enumerate(product_urls, 1):
        try:
            driver.get(url)
            try:
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div.pro-name')))
            except TimeoutException:
                print(f"  Timeout loading {url}. Skipping.")
                continue

            product_name = safe_find_text(driver, By.CSS_SELECTOR, 'div.pro-name')
            print(f"  Scraping {i}/{total_products} -> {product_name}")

            product_data = {
                'Category': category_breadcrumb_str,
                'Product URL': url,
                'Product Name': product_name,
                'SKU': safe_find_text(driver, By.CSS_SELECTOR, 'div.pro-sku .psku-spa'),
                'Price': safe_find_text(driver, By.CSS_SELECTOR, 'span.price-wrapper'),
                'Brand': 'Interlude Home',
                'Tearsheet': safe_find_attr(driver, By.CSS_SELECTOR, 'div.pdp_tear a.configtear', 'href'),
                'Dimensions': safe_find_text(driver, By.CSS_SELECTOR, '.pro-dimension .spec-head.prdimens'),
                'Finish': safe_find_text(driver, By.CSS_SELECTOR, '.pro-finish .spec-head.prfinish'),
                'Material': safe_find_text(driver, By.CSS_SELECTOR, '.pro-material .spec-head.prmaterial'),
                'Unit': safe_find_text(driver, By.CSS_SELECTOR, '.pro-unit .spec-head.prunit')
            }
            
            desc = safe_find_text(driver, By.CSS_SELECTOR, '.product.attribute.description .value')
            disc = safe_find_text(driver, By.CSS_SELECTOR, '.product.attribute.description .disclamier .value')
            product_data['Description'] = f"{desc}\n{disc}".strip()
            product_data['Full Description'] = safe_find_attr(driver, By.CSS_SELECTOR, '#description[data-role="content"]', 'innerHTML')

            # --- SEPARATED IMAGE EXTRACTION ---
            thumb_urls, main_urls = extract_images_separated(driver)

            # 1. Assign Thumbnails (Image_1 to Image_4)
            # Use steps to get different angles from 360 view if possible (e.g., indices 0, 9, 18, 27)
            # Otherwise sequential
            if len(thumb_urls) > 10: # Likely a 360 slider
                indices = [0, 9, 18, 27] # Approximate 90 degree turns
                for k, idx in enumerate(indices):
                     key = f'Image_{k+1}'
                     if idx < len(thumb_urls):
                         product_data[key] = thumb_urls[idx]
                     else:
                         product_data[key] = 'N/A'
            else:
                for k in range(4):
                    product_data[f'Image_{k+1}'] = thumb_urls[k] if k < len(thumb_urls) else 'N/A'
            
            # 2. Assign Main Images (Main_Image_1 to Main_Image_3)
            for k in range(3):
                product_data[f'Main_Image_{k+1}'] = main_urls[k] if k < len(main_urls) else 'N/A'

            all_data_list.append(product_data)
            
            if len(all_data_list) % PRODUCTS_PER_AUTOSAVE == 0:
                save_to_excel(all_data_list, filename, is_autosave=True)

        except Exception as e:
            print(f"    Error on {url}: {e}")

# --- 4. Main Execution ---
if __name__ == "__main__":
    all_products_details = [] 
    driver = get_driver()
    
    try:
        for category in CATEGORIES_TO_SCRAPE:
            product_urls, breadcrumbs = get_product_urls_from_category(driver, category['name'], category['url'])
            if product_urls:
                scrape_product_details(driver, breadcrumbs, product_urls, all_products_details, EXCEL_FILENAME)
            
    except KeyboardInterrupt:
        print("\n--- Stopped by User ---")
    finally:
        save_to_excel(all_products_details, EXCEL_FILENAME, is_autosave=False)
        driver.quit()
        sys.exit(0)