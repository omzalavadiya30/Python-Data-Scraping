import time
import pandas as pd
import signal
import sys
import urllib.parse
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# --- GLOBAL VARIABLES FOR SAVING DATA ---
ALL_DATA = []
OUTPUT_FILE = "BrownJordan_Complete_Data.xlsx"

# --- 1. FULL CATEGORY LIST (Snippet - Include your full list here) ---
CATEGORIES = [
    # --- TYPE: SEATING ---
    {"Category Type": "Type", "Category Name": "Arm Chairs", "Category URL": "https://www.brownjordan.com/products/type/seating/arm-chairs"},
    {"Category Type": "Type", "Category Name": "Bar Stools", "Category URL": "https://www.brownjordan.com/products/type/seating/bar-stools"},
    {"Category Type": "Type", "Category Name": "Benches", "Category URL": "https://www.brownjordan.com/products/type/seating/benches"},
    {"Category Type": "Type", "Category Name": "Chaise Lounges", "Category URL": "https://www.brownjordan.com/products/type/seating/chaise-lounges"},
    {"Category Type": "Type", "Category Name": "Daybeds", "Category URL": "https://www.brownjordan.com/products/type/seating/daybeds"},
    {"Category Type": "Type", "Category Name": "Dining Chairs", "Category URL": "https://www.brownjordan.com/products/type/seating/dining-chairs"},
    {"Category Type": "Type", "Category Name": "Lounge Chairs", "Category URL": "https://www.brownjordan.com/products/type/seating/lounge-chairs"},
    {"Category Type": "Type", "Category Name": "Loveseats", "Category URL": "https://www.brownjordan.com/products/type/seating/loveseats"},
    {"Category Type": "Type", "Category Name": "Ottomans", "Category URL": "https://www.brownjordan.com/products/type/seating/ottomans"},
    {"Category Type": "Type", "Category Name": "Sand Chairs", "Category URL": "https://www.brownjordan.com/products/type/seating/sand-chairs"},
    {"Category Type": "Type", "Category Name": "Sectionals", "Category URL": "https://www.brownjordan.com/products/type/seating/sectionals"},
    {"Category Type": "Type", "Category Name": "Sofas", "Category URL": "https://www.brownjordan.com/products/type/seating/sofas"},
    {"Category Type": "Type", "Category Name": "Swivel Rockers", "Category URL": "https://www.brownjordan.com/products/type/seating/swivel-rockers"},

    # --- TYPE: TABLES ---
    {"Category Type": "Type", "Category Name": "Bar Tables", "Category URL": "https://www.brownjordan.com/products/type/tables/bar-tables"},
    {"Category Type": "Type", "Category Name": "Chat Tables", "Category URL": "https://www.brownjordan.com/products/type/tables/chat-tables"},
    {"Category Type": "Type", "Category Name": "Coffee Tables", "Category URL": "https://www.brownjordan.com/products/type/tables/coffee-tables"},
    {"Category Type": "Type", "Category Name": "Dining Tables", "Category URL": "https://www.brownjordan.com/products/type/tables/dining-tables"},
    {"Category Type": "Type", "Category Name": "Occasional Tables", "Category URL": "https://www.brownjordan.com/products/type/tables/occasional-tables"},
    {"Category Type": "Type", "Category Name": "Table Bases", "Category URL": "https://www.brownjordan.com/products/type/tables/table-bases"},

    # --- TYPE: ACCESSORIES ---
    {"Category Type": "Type", "Category Name": "Bar Carts", "Category URL": "https://www.brownjordan.com/products/type/accessories/bar-carts"},
    {"Category Type": "Type", "Category Name": "Covers", "Category URL": "https://www.brownjordan.com/products/type/accessories/covers"},
    {"Category Type": "Type", "Category Name": "Fire Tables", "Category URL": "https://www.brownjordan.com/products/type/accessories/fire-tables"},
    {"Category Type": "Type", "Category Name": "Lanterns", "Category URL": "https://www.brownjordan.com/products/type/accessories/lanterns"},
    {"Category Type": "Type", "Category Name": "Pillows", "Category URL": "https://www.brownjordan.com/products/type/accessories/pillows"},
    {"Category Type": "Type", "Category Name": "Planters", "Category URL": "https://www.brownjordan.com/products/type/accessories/planters"},
    {"Category Type": "Type", "Category Name": "Poufs", "Category URL": "https://www.brownjordan.com/products/type/accessories/poufs"},
    {"Category Type": "Type", "Category Name": "Rugs", "Category URL": "https://www.brownjordan.com/products/type/accessories/rugs"},
    {"Category Type": "Type", "Category Name": "Trays", "Category URL": "https://www.brownjordan.com/products/type/accessories/trays"},
    {"Category Type": "Type", "Category Name": "Umbrella Bases", "Category URL": "https://www.brownjordan.com/products/type/accessories/umbrella-bases"},
    {"Category Type": "Type", "Category Name": "Umbrellas", "Category URL": "https://www.brownjordan.com/products/type/accessories/umbrellas"},

    # --- COLLECTIONS ---
    {"Category Type": "Collection", "Category Name": "Adapt", "Category URL": "https://www.brownjordan.com/products/collection/tables-accessories/adapt"},
    {"Category Type": "Collection", "Category Name": "Atlas", "Category URL": "https://www.brownjordan.com/products/collection/tables-accessories/atlas"},
    {"Category Type": "Collection", "Category Name": "Calcutta", "Category URL": "https://www.brownjordan.com/products/collection/cushion/calcutta"},
    {"Category Type": "Collection", "Category Name": "Faro", "Category URL": "https://www.brownjordan.com/products/collection/rope/faro"},
    {"Category Type": "Collection", "Category Name": "Flight Sling", "Category URL": "https://www.brownjordan.com/products/collection/sling/flight-sling"},
    {"Category Type": "Collection", "Category Name": "Fremont", "Category URL": "https://www.brownjordan.com/products/collection/multi-type/fremont"},
    {"Category Type": "Collection", "Category Name": "Fusion", "Category URL": "https://www.brownjordan.com/products/collection/resinweave/fusion"},
    {"Category Type": "Collection", "Category Name": "H", "Category URL": "https://www.brownjordan.com/products/collection/rope/h"},
    {"Category Type": "Collection", "Category Name": "Huntley", "Category URL": "https://www.brownjordan.com/products/collection/cushion/huntley"},
    {"Category Type": "Collection", "Category Name": "Juno", "Category URL": "https://www.brownjordan.com/products/collection/tables-accessories/juno"},
    {"Category Type": "Collection", "Category Name": "Kantan", "Category URL": "https://www.brownjordan.com/products/collection/multi-type/kantan"},
    {"Category Type": "Collection", "Category Name": "Lisbon", "Category URL": "https://www.brownjordan.com/products/collection/rope/lisbon"},
    {"Category Type": "Collection", "Category Name": "Luca", "Category URL": "https://www.brownjordan.com/products/collection/cushion/luca"},
    {"Category Type": "Collection", "Category Name": "Moto", "Category URL": "https://www.brownjordan.com/products/collection/cushion/moto"},
    {"Category Type": "Collection", "Category Name": "Oliver", "Category URL": "https://www.brownjordan.com/products/collection/rope/oliver"},
    {"Category Type": "Collection", "Category Name": "Oscar", "Category URL": "https://www.brownjordan.com/products/collection/rope/oscar"},
    {"Category Type": "Collection", "Category Name": "Oscar II", "Category URL": "https://www.brownjordan.com/products/collection/strap/oscar-ii"},
    {"Category Type": "Collection", "Category Name": "Parkway", "Category URL": "https://www.brownjordan.com/products/collection/multi-type/parkway"},
    {"Category Type": "Collection", "Category Name": "Pasadena", "Category URL": "https://www.brownjordan.com/products/collection/multi-type/pasadena"},
    {"Category Type": "Collection", "Category Name": "Softscape", "Category URL": "https://www.brownjordan.com/products/collection/multi-type/softscape"},
    {"Category Type": "Collection", "Category Name": "Solenne", "Category URL": "https://www.brownjordan.com/products/collection/cushion/solenne"},
    {"Category Type": "Collection", "Category Name": "Sol Y Luna", "Category URL": "https://www.brownjordan.com/products/collection/cushion/sol-y-luna"},
    {"Category Type": "Collection", "Category Name": "Southampton", "Category URL": "https://www.brownjordan.com/products/collection/resinweave/southampton"},
    {"Category Type": "Collection", "Category Name": "Still", "Category URL": "https://www.brownjordan.com/products/collection/cushion/still"},
    {"Category Type": "Collection", "Category Name": "Stretch", "Category URL": "https://www.brownjordan.com/products/collection/strap/stretch"},
    {"Category Type": "Collection", "Category Name": "Swim", "Category URL": "https://www.brownjordan.com/products/collection/sling/swim"},
    {"Category Type": "Collection", "Category Name": "Summit Umbrellas", "Category URL": "https://www.brownjordan.com/products/collection/tables-accessories/summit-umbrellas"},
    {"Category Type": "Collection", "Category Name": "Trentino", "Category URL": "https://www.brownjordan.com/products/collection/cushion/trentino"},
    {"Category Type": "Collection", "Category Name": "Venetian", "Category URL": "https://www.brownjordan.com/products/collection/cushion/venetian"},
    {"Category Type": "Collection", "Category Name": "Walter Lamb", "Category URL": "https://www.brownjordan.com/products/collection/rope/walter-lamb"},

    # --- DIRECT LINKS ---
    {"Category Type": "Direct", "Category Name": "New Releases", "Category URL": "https://www.brownjordan.com/new-releases"},
    {"Category Type": "Direct", "Category Name": "Last Chance", "Category URL": "https://www.brownjordan.com/sale"},

    # --- MATERIALS (Textiles) ---
    {"Category Type": "Materials", "Category Name": "Suncloth Fabric", "Category URL": "https://www.brownjordan.com/materials/textiles/suncloth-fabric"},
    {"Category Type": "Materials", "Category Name": "Brown Jordan x Kravet", "Category URL": "https://www.brownjordan.com/materials/textiles/brown-jordan-x-kravet"},
    {"Category Type": "Materials", "Category Name": "Versatex Sling", "Category URL": "https://www.brownjordan.com/materials/textiles/versatex-sling"},
    {"Category Type": "Materials", "Category Name": "Rope", "Category URL": "https://www.brownjordan.com/materials/textiles/rope"},
    {"Category Type": "Materials", "Category Name": "Suncloth Strap", "Category URL": "https://www.brownjordan.com/materials/textiles/suncloth-strap"},
    {"Category Type": "Materials", "Category Name": "Suncloth Lace", "Category URL": "https://www.brownjordan.com/materials/textiles/suncloth-lace"},
    {"Category Type": "Materials", "Category Name": "Rugs (Material)", "Category URL": "https://www.brownjordan.com/materials/rugs"},

    # --- MATERIALS (Non-Textiles) ---
    {"Category Type": "Materials", "Category Name": "Frame Finishes", "Category URL": "https://www.brownjordan.com/materials/non-textiles/frame-finishes"},
    {"Category Type": "Materials", "Category Name": "Vinyl Lace Straps", "Category URL": "https://www.brownjordan.com/materials/non-textiles/vinyl-lace-straps"},
    {"Category Type": "Materials", "Category Name": "Fiber", "Category URL": "https://www.brownjordan.com/materials/non-textiles/fiber"},
    {"Category Type": "Materials", "Category Name": "Teak", "Category URL": "https://www.brownjordan.com/materials/non-textiles/teak"},
    {"Category Type": "Materials", "Category Name": "Ceramic", "Category URL": "https://www.brownjordan.com/materials/non-textiles/ceramic"},
    {"Category Type": "Materials", "Category Name": "Dekton", "Category URL": "https://www.brownjordan.com/materials/non-textiles/dekton"},
    {"Category Type": "Materials", "Category Name": "Glass and Acrylic", "Category URL": "https://www.brownjordan.com/materials/non-textiles/glass-and-acrylic"},
    {"Category Type": "Materials", "Category Name": "Aluminum Tops", "Category URL": "https://www.brownjordan.com/materials/non-textiles/aluminum-tops"},
    {"Category Type": "Materials", "Category Name": "Vector Tops", "Category URL": "https://www.brownjordan.com/materials/non-textiles/vector-tops"},
    {"Category Type": "Materials", "Category Name": "Fire Media", "Category URL": "https://www.brownjordan.com/materials/non-textiles/fire-media"},
]

# --- 2. SETUP DRIVER ---
def setup_driver():
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    # options.add_argument("--headless") 
    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(300) 
    return driver

# --- 3. URL CLEANER FUNCTION (UPDATED) ---
def clean_image_url(raw_url):
    """
    Cleans and standardizes image URLs.
    1. Passes direct Cylindo URLs through.
    2. Decodes Next.js optimized URLs to find the source.
    3. Fixes relative paths.
    """
    if not raw_url:
        return "N/A"
    
    final_url = raw_url.strip()

    # CASE A: Next.js Optimized URL (contains /_next/image?url=...)
    if "/_next/image" in final_url and "url=" in final_url:
        try:
            parsed = urllib.parse.urlparse(final_url)
            query_params = urllib.parse.parse_qs(parsed.query)
            # Extract the real 'url' param
            extracted_url = query_params.get('url', [None])[0]
            if extracted_url:
                final_url = extracted_url
        except:
            pass # Keep original if parsing fails

    # CASE B: Handle Protocol-Relative URLs (starts with //)
    if final_url.startswith("//"):
        final_url = "https:" + final_url

    # CASE C: Handle Root-Relative URLs (starts with / but not //)
    elif final_url.startswith("/") and not final_url.startswith("//"):
        final_url = "https://www.brownjordan.com" + final_url

    # CASE D: Cylindo URLs are usually absolute, but if missing https...
    if "cylindo.com" in final_url and not final_url.startswith("http"):
        final_url = "https://" + final_url.lstrip("/")

    return final_url

# --- 4. SAVE HELPER ---
def save_data():
    if ALL_DATA:
        print(f"\n[SYSTEM] Saving {len(ALL_DATA)} rows to '{OUTPUT_FILE}'...")
        df = pd.DataFrame(ALL_DATA)
        cols = [
            'Category', 'Product URL', 'Product Name', 'SKU', 'Price',
            'Brand', 'Description', 
            'Arm Height', 'Height', 'Seat Height', 'Width', 'Length', 'Seat Depth',
            'Grade', 'Construction', 'Fabric Type', 'Use',
            'Image 1', 'Image 2', 'Image 3', 'Image 4'
        ]
        existing_cols = [c for c in cols if c in df.columns]
        remaining_cols = [c for c in df.columns if c not in cols]
        df = df[existing_cols + remaining_cols]
        df.to_excel(OUTPUT_FILE, index=False)
        print("[SYSTEM] Save Complete.")

def signal_handler(sig, frame):
    print("\n\n[SYSTEM] Interrupt received (Ctrl+C). Saving data before exiting...")
    save_data()
    sys.exit(0)

signal.signal(signal.SIGINT, signal_handler)

# --- 5. CATEGORY SCRAPER ---
def get_category_data(driver, category):
    url = category["Category URL"]
    print(f"\n--- Processing Category: {category['Category Name']} ---")
    driver.get(url)
    time.sleep(4) 

    # Extract Breadcrumbs (Path)
    category_path_str = "Home > Products"
    try:
        breadcrumb_items = driver.find_elements(By.XPATH, "//nav[contains(@class, 'chakra-breadcrumb')]//li")
        if breadcrumb_items:
            cats = [item.text.replace("/", "").strip() for item in breadcrumb_items]
            category_path_str = " > ".join([c for c in cats if c])
    except: pass

    # Infinite Scroll to get all URLs
    product_urls = set()
    last_count = 0
    no_change_passes = 0
    
    while True:
        format1 = driver.find_elements(By.XPATH, "//div[contains(@class, 'chakra-linkbox')]")
        format2 = driver.find_elements(By.XPATH, "//div[@data-sentry-component='ProductCardConfigurable']")
        current_count = len(format1) + len(format2)
        
        try:
            spinner = driver.find_element(By.XPATH, "//div[contains(@class, 'css-1cs38w7')] | //div[contains(@class, 'chakra-spinner')]")
            driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", spinner)
            loader_found = True
        except NoSuchElementException:
            all_products = format1 + format2
            if len(all_products) > 0:
                driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", all_products[-1])
            else:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            loader_found = False

        time.sleep(3) 

        format1_new = driver.find_elements(By.XPATH, "//div[contains(@class, 'chakra-linkbox')]")
        format2_new = driver.find_elements(By.XPATH, "//div[@data-sentry-component='ProductCardConfigurable']")
        new_count = len(format1_new) + len(format2_new)

        if new_count == last_count:
            if not loader_found:
                break
            no_change_passes += 1
            driver.execute_script("window.scrollBy(0, -300);")
            time.sleep(1)
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            if no_change_passes >= 3:
                break
        else:
            no_change_passes = 0
        last_count = new_count

    # Collect URLs
    for card in driver.find_elements(By.XPATH, "//div[contains(@class, 'chakra-linkbox')]"):
        try:
            link = card.find_element(By.XPATH, ".//a[contains(@class, 'chakra-linkbox__overlay')]")
            product_urls.add(link.get_attribute("href"))
        except: pass
    
    for card in driver.find_elements(By.XPATH, "//div[@data-sentry-component='ProductCardConfigurable']"):
        try:
            link = card.find_element(By.XPATH, ".//a[h4]")
            product_urls.add(link.get_attribute("href"))
        except: pass

    print(f"   > URLs Found: {len(product_urls)}")
    return list(product_urls), category_path_str

# --- 6. PRODUCT DETAILS EXTRACTOR ---
def scrape_product_details(driver, product_url, category_path, current_idx, total_idx):
    driver.get(product_url)
    
    data = {
        "Category": category_path,
        "Product URL": product_url,
        "Brand": "Brown Jordan",
        "Product Name": "N/A",
        "SKU": "N/A",
        "Price": "N/A",
        "Description": "N/A",
        "Arm Height": "N/A", "Height": "N/A", "Seat Height": "N/A", 
        "Width": "N/A", "Length": "N/A", "Seat Depth": "N/A",
        "Grade": "N/A", "Construction": "N/A", "Fabric Type": "N/A", "Use": "N/A",
        "Image 1": "N/A", "Image 2": "N/A", "Image 3": "N/A", "Image 4": "N/A"
    }

    try:
        # Basic Info
        try:
            name_el = driver.find_element(By.XPATH, "//h1[contains(@class, 'chakra-heading')]")
            data["Product Name"] = name_el.text.strip()
        except: pass

        try:
            sku_el = driver.find_element(By.XPATH, "//span[contains(text(), 'SKU')]/parent::div")
            data["SKU"] = sku_el.text.replace("SKU", "").strip()
        except: pass

        try:
            price_el = driver.find_element(By.XPATH, "//div[contains(@class, 'chakra-stack')]//p[contains(text(), '$')]")
            data["Price"] = price_el.text.strip()
        except: pass

        try:
            desc_el = driver.find_element(By.XPATH, "//div[contains(@class, 'product-description')]")
            data["Description"] = desc_el.text.strip()
        except: pass

        # Specs (Dimensions & Details)
        spec_headers = driver.find_elements(By.XPATH, "//div[contains(@class, 'css-1qb6pms')]")
        for header in spec_headers:
            header_text = header.text.strip().lower()
            try:
                ul_element = header.find_element(By.XPATH, "./following-sibling::ul")
                items = ul_element.find_elements(By.TAG_NAME, "li")
                for li in items:
                    raw_text = li.text.strip()
                    if "dimensions" in header_text:
                        parts = raw_text.split(" ", 1)
                        if len(parts) == 2:
                            key, val = parts[0].strip(), parts[1].strip()
                            if key == "AH": data["Arm Height"] = val
                            elif key == "H": data["Height"] = val
                            elif key == "SH": data["Seat Height"] = val
                            elif key == "W": data["Width"] = val
                            elif key == "L": data["Length"] = val
                            elif key == "SD": data["Seat Depth"] = val
                    elif "details" in header_text:
                        if ":" in raw_text:
                            key, val = raw_text.split(":", 1)
                            key = key.strip().lower()
                            val = val.strip()
                            if key == "grade": data["Grade"] = val
                            elif key == "construction": data["Construction"] = val
                            elif key == "fabric type": data["Fabric Type"] = val
                            elif key == "use": data["Use"] = val
            except: continue

        # --- IMAGE EXTRACTION (with Cleaning) ---
        images_found = []
        
        # 1. Cylindo (Format 1)
        if not images_found:
            cylindo_imgs = driver.find_elements(By.XPATH, "//img[contains(@src, 'cylindo')]")
            for img in cylindo_imgs[:4]:
                src = clean_image_url(img.get_attribute("src"))
                if src and src not in images_found:
                    images_found.append(src)

        # 2. Carousel Thumbnails (Format 2 - Rugs)
        if not images_found:
            thumbnails = driver.find_elements(By.XPATH, "//div[contains(@class, 'css-oj0i39')]//img")
            for thumb in thumbnails:
                # Try srcset first for higher res
                srcset = thumb.get_attribute("srcset")
                src = thumb.get_attribute("src")
                
                best_src = src
                if srcset:
                    try:
                        # Grab last url in srcset
                        best_src = srcset.split(",")[-1].strip().split(" ")[0]
                    except: pass
                
                clean_src = clean_image_url(best_src)
                if clean_src and clean_src not in images_found:
                    images_found.append(clean_src)
            
            # If still no thumbnails, try the main carousel image
            if not images_found:
                try:
                    main_img = driver.find_element(By.XPATH, "//div[contains(@class, 'carousel__inner-slide')]//img")
                    clean_src = clean_image_url(main_img.get_attribute("src"))
                    images_found.append(clean_src)
                except: pass

        # 3. Standard Single Image (Format 3)
        if not images_found:
            try:
                single_img = driver.find_element(By.XPATH, "//div[contains(@class, 'chakra-aspect-ratio')]//img")
                # Try srcset then src
                srcset = single_img.get_attribute("srcset")
                src = single_img.get_attribute("src")
                
                best_src = src
                if srcset:
                    try:
                        best_src = srcset.split(",")[-1].strip().split(" ")[0]
                    except: pass
                
                clean_src = clean_image_url(best_src)
                if clean_src:
                    images_found.append(clean_src)
            except: pass

        # Assign to dict
        for i in range(4):
            if i < len(images_found):
                data[f"Image {i+1}"] = images_found[i]

    except Exception: pass

    print(f"Scraping {current_idx}/{total_idx} -> {data['Product Name']} -> {data['SKU']}")
    return data

# --- 7. MAIN EXECUTION ---
def main():
    driver = setup_driver()
    try:
        total_cats = len(CATEGORIES)
        for i, cat in enumerate(CATEGORIES):
            print(f"[{i+1}/{total_cats}] Starting category...")
            for attempt in range(3):
                try:
                    product_urls, cat_path = get_category_data(driver, cat)
                    total_products = len(product_urls)
                    print(f"   > Entering details loop for {total_products} products...")
                    
                    for j, p_url in enumerate(product_urls):
                        product_data = scrape_product_details(driver, p_url, cat_path, j+1, total_products)
                        ALL_DATA.append(product_data)
                        if len(ALL_DATA) % 50 == 0:
                            save_data()
                    break 
                except Exception as e:
                    print(f"   Error: {e}. Retrying...")
                    driver.refresh()
                    time.sleep(5)
        save_data() 
    except Exception as e:
        print(f"Critical Error: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()