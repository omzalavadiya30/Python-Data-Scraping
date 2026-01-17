import pandas as pd
import time
import sys
import re
import signal
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager

# --- CONFIGURATION ---
EXCEL_FILENAME = 'VisualComfort_Final_Data.xlsx'
AUTOSAVE_INTERVAL = 50

# --- GLOBAL STOP FLAG ---
STOP_EXECUTION = False

def signal_handler(sig, frame):
    """Sets the stop flag to exit loops immediately."""
    global STOP_EXECUTION
    print("\n\n[!!!] Ctrl+C Detected! Stopping gracefully... (Data will be saved)")
    STOP_EXECUTION = True

# Register the signal handler
signal.signal(signal.SIGINT, signal_handler)

# --- URLS ---
target_map = {
    "Ceiling": [
        "https://www.visualcomfort.com/us/c/ceiling",
        "https://www.visualcomfort.com/us/c/ceiling/new-introductions",
        "https://www.visualcomfort.com/us/c/ceiling/chandelier",
        "https://www.visualcomfort.com/us/c/ceiling/flush-mount",
        "https://www.visualcomfort.com/us/c/ceiling/pendant",
        "https://www.visualcomfort.com/us/c/ceiling/lantern",
        "https://www.visualcomfort.com/us/c/ceiling/hanging-shade",
        "https://www.visualcomfort.com/us/c/ceiling/linear",
        "https://www.visualcomfort.com/us/c/bulbs",
        "https://www.visualcomfort.com/us/c/sale?Category=Ceiling_-_Chandelier%2CCeiling_-_Flush_Mount%2CCeiling_-_Hanging_Shade%2CCeiling_-_Lantern%2CCeiling_-_Linear%2CCeiling_-_Pendant"
    ],
    "Wall": [
        "https://www.visualcomfort.com/us/c/wall",
        "https://www.visualcomfort.com/us/c/wall/new-introductions",
        "https://www.visualcomfort.com/us/c/wall/decorative-wall",
        "https://www.visualcomfort.com/us/c/wall/bath",
        "https://www.visualcomfort.com/us/c/wall/task",
        "https://www.visualcomfort.com/us/c/wall/picture",
        "https://www.visualcomfort.com/us/c/wall/mirrors",
        "https://www.visualcomfort.com/us/c/sale?Category=Wall_-_Bath%2CWall_-_Decorative%2CWall_-_Picture_Lights%2CWall_-_Task"
    ],
    "Table": [
        "https://www.visualcomfort.com/us/c/table",
        "https://www.visualcomfort.com/us/c/table/new-introductions",
        "https://www.visualcomfort.com/us/c/table/decorative",
        "https://www.visualcomfort.com/us/c/table/table-task",
        "https://www.visualcomfort.com/us/c/table/cordless-and-rechargeable",
        "https://www.visualcomfort.com/us/c/sale?Category=Table_-_Decorative%2CTable_-_Task"
    ],
    "Floor": [
        "https://www.visualcomfort.com/us/c/floor",
        "https://www.visualcomfort.com/us/c/floor/new-introductions",
        "https://www.visualcomfort.com/us/c/floor/decorative",
        "https://www.visualcomfort.com/us/c/floor/task",
        "https://www.visualcomfort.com/us/c/sale?Category=Floor_-_Decorative%2CFloor_-_Task"
    ],
    "Outdoor": [
        "https://www.visualcomfort.com/us/c/outdoor",
        "https://www.visualcomfort.com/us/c/outdoor/new-introductions",
        "https://www.visualcomfort.com/us/c/outdoor/wall",
        "https://www.visualcomfort.com/us/c/outdoor/ceiling",
        "https://www.visualcomfort.com/us/c/outdoor/table-and-floor",
        "https://www.visualcomfort.com/us/c/outdoor/outdoor-rechargeable-portables",
        "https://www.visualcomfort.com/us/c/outdoor/post",
        "https://www.visualcomfort.com/us/c/outdoor/gas-lanterns",
        "https://www.visualcomfort.com/us/c/sale?Category=Outdoor_-_Bollard_%26_Path%2COutdoor_-_Ceiling%2COutdoor_-_Post%2COutdoor_-_Wall"
    ],
    "Fans": [
        "https://www.visualcomfort.com/us/c/fans",
        "https://www.visualcomfort.com/us/c/fans/new-fan-introductions",
        "https://www.visualcomfort.com/us/c/fans/indoor",
        "https://www.visualcomfort.com/us/c/fans/indoor-outdoor",
        "https://www.visualcomfort.com/us/c/fans/accessories",
        "https://www.visualcomfort.com/us/c/sale?Category=Fans_-_Indoor%2CFans_-_Indoor%2FOutdoor"
    ],
    "Collections": [
        "https://www.visualcomfort.com/us/c/our-collections/signature-collection",
        "https://www.visualcomfort.com/us/c/our-collections/modern-collection",
        "https://www.visualcomfort.com/us/c/our-collections/studio-collection",
        "https://www.visualcomfort.com/us/c/our-collections/fan-collection",
        "https://www.visualcomfort.com/us/c/our-collections/generation-lighting"
    ]
}

# --- HELPER FUNCTIONS ---

def get_driver():
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    # options.add_argument('--headless') 
    options.page_load_strategy = 'eager' 
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def check_popup(driver):
    if STOP_EXECUTION: return 
    try:
        close_btns = driver.find_elements(By.CSS_SELECTOR, "div.popup-modal span.close")
        for btn in close_btns:
            if btn.is_displayed():
                driver.execute_script("arguments[0].click();", btn)
                time.sleep(1)
    except: pass

def save_to_excel(data, filename):
    if not data: return
    print(f"\n[SYSTEM] Saving {len(data)} records to {filename}...")
    try:
        df = pd.DataFrame(data)
        if 'Product URL' in df.columns:
            df.drop_duplicates(subset=['Product URL'], inplace=True)
        df.to_excel(filename, index=False)
        print("[SYSTEM] Save Complete.")
    except Exception as e:
        print(f"[ERROR] Save failed: {e}")

def safe_text(driver, selector):
    try: return driver.find_element(By.CSS_SELECTOR, selector).text.strip()
    except: return "N/A"

def safe_attr(driver, selector, attr):
    try: return driver.find_element(By.CSS_SELECTOR, selector).get_attribute(attr)
    except: return "N/A"

# --- EXTRACTION LOGIC ---

def extract_product_details(driver, product_url, main_cat):
    # 1. Stop Check before load
    if STOP_EXECUTION: return None

    try:
        driver.get(product_url)
    except: return None

    # 2. Stop Check before wait
    if STOP_EXECUTION: return None

    wait = WebDriverWait(driver, 15)
    
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'h1.page-title')))
    except: return None

    check_popup(driver)
    item = {}

    # --- 3. SMART SCROLLING SEQUENCE (Tips -> Visuals -> Top) ---
    if not STOP_EXECUTION:
        try:
            # A. Scroll to TIPS (Slowly)
            tips_el = driver.find_elements(By.CSS_SELECTOR, ".tips-slider, .tips-item")
            if tips_el:
                driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", tips_el[0])
                time.sleep(2.5) # Wait for render
            
            if STOP_EXECUTION: return None

            # B. Scroll to VISUAL IMAGES (Slowly)
            visual_el = driver.find_elements(By.CSS_SELECTOR, ".olapic-slider-wrapper")
            if visual_el:
                driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", visual_el[0])
                time.sleep(2.5) # Wait for render

            if STOP_EXECUTION: return None

            # C. Scroll back to TOP for basic info
            driver.execute_script("window.scrollTo(0, 0);")
            time.sleep(0.5)
        except: pass

    # --- 4. DATA EXTRACTION ---

    # Breadcrumbs (Text Extraction)
    try:
        crumbs = driver.find_elements(By.CSS_SELECTOR, "div.breadcrumbs ul.items li.item")
        path = [c.text.strip() for c in crumbs if c.text.strip()]
        item['Category'] = " / ".join(path) if path else "N/A"
    except: item['Category'] = "N/A"

    # Basic Info
    item['Product URL'] = product_url
    item['Product Name'] = safe_text(driver, 'h1.page-title span')
    item['SKU'] = safe_text(driver, 'div[itemprop="sku"]')
    item['Brand']= "Visual Comfort"
    item['Price'] = safe_text(driver, 'span.price')
    item['Designer'] = safe_text(driver, 'div.product.attribute.designer .value')

    # Descriptions
    item['Description'] = safe_text(driver, 'div.additional-description .content')
    try:
        item['Full Description HTML'] = driver.find_element(By.CSS_SELECTOR, 'div.product-info-main').get_attribute('outerHTML')
    except: item['Full Description HTML'] = "N/A"

    # Tips HTML (Extracted after scroll)
    try:
        tips_node = driver.find_elements(By.CSS_SELECTOR, ".tips-slider, .tips-item")
        item['Tips HTML'] = tips_node[0].get_attribute('outerHTML') if tips_node else "N/A"
    except: item['Tips HTML'] = "N/A"

    # PDFs
    item['Spec Sheet PDF'] = safe_attr(driver, 'a[href*="spec_sheet"]', 'href')
    item['Install Guide PDF'] = safe_attr(driver, 'a[href*="install_guide"]', 'href')

    # --- IMAGES (1-6) ---
    # Image 1 (Main Stage)
    try:
        img1 = safe_attr(driver, '.fotorama__stage__frame img', 'src')
        item['Image1'] = img1 if img1 != "N/A" else ""
    except: pass

    # Images 2-6 (Thumbnails)
    try:
        thumbs = driver.find_elements(By.CSS_SELECTOR, '.fotorama__nav__frame img')
        for idx, thumb in enumerate(thumbs):
            if idx < 5: # We already have Img1, so we get up to 5 more (Img2-Img6)
                item[f'Image{idx+2}'] = thumb.get_attribute('src')
    except: pass

    # --- VISUAL IMAGES (1-3) ---
    try:
        # Target the specific LI elements in the Olapic slider
        visual_lis = driver.find_elements(By.CSS_SELECTOR, ".olapic-slider-wrapper li.instagram_graph, .olapic-slider-wrapper li.instagram")
        
        v_count = 1
        for li in visual_lis:
            if v_count > 3: break
            
            # Priority 1: data-src attribute
            src = li.get_attribute('data-src')
            
            # Priority 2: Background Image Style regex
            if not src:
                style = li.get_attribute('style')
                if style and 'url(' in style:
                    match = re.search(r'url\(["\']?(.*?)["\']?\)', style)
                    if match:
                        src = match.group(1)
            
            if src:
                item[f'VisualImage{v_count}'] = src
                v_count += 1
    except: pass

    # --- ATTRIBUTES ---
    # Initialize empty
    attr_keys = [
        'Attr_Height', 'Attr_Weight', 'Attr_Width', 'Attr_Length', 'Attr_Canopy', 
        'Attr_Socket', 'Attr_Wattage', 'Attr_Diameter', 'Attr_Motor_Size', 
        'Attr_LightSource', 'Attr_Fixture_Height', 'Attr_Min_Custom_Height', 
        'Attr_OA_Height', 'Attr_Chain_Length', 'Attr_Shade_Details', 'Attr_Base', 
        'Attr_Extension', 'Attr_Blackplate', 'Attr_Rating', 'Attr_CFM_Average', 
        'Attr_CFM_Low', 'Attr_CFM_High', 'Attr_CFM_Wattage_Low', 'Attr_CFM_Wattage_High'
    ]
    for k in attr_keys: item[k] = "N/A"

    try:
        rows = driver.find_elements(By.CSS_SELECTOR, '#spec-inch-tab table.data-table.product-attribute-specs-table tr')
        for row in rows:
            txt = row.text.strip()
            # Simple mapping based on startswith
            if txt.startswith('Height:'): item['Attr_Height'] = txt.replace('Height:', '').strip()
            elif txt.startswith('Weight:'): item['Attr_Weight'] = txt.replace('Weight:', '').strip()
            elif txt.startswith('Width:'): item['Attr_Width'] = txt.replace('Width:', '').strip()
            elif txt.startswith('Length:'): item['Attr_Length'] = txt.replace('Length:', '').strip()
            elif txt.startswith('Canopy:'): item['Attr_Canopy'] = txt.replace('Canopy:', '').strip()
            elif txt.startswith('Socket:'): item['Attr_Socket'] = txt.replace('Socket:', '').strip()
            elif txt.startswith('Wattage:'): item['Attr_Wattage'] = txt.replace('Wattage:', '').strip()
            elif txt.startswith('Diameter:'): item['Attr_Diameter'] = txt.replace('Diameter:', '').strip()
            elif txt.startswith('Motor Size:'): item['Attr_Motor_Size'] = txt.replace('Motor Size:', '').strip()
            elif txt.startswith('Lightsource:'): item['Attr_LightSource'] = txt.replace('Lightsource:', '').strip()
            elif txt.startswith('Fixture Height:'): item['Attr_Fixture_Height'] = txt.replace('Fixture Height:', '').strip()
            elif txt.startswith('Min. Custom Height:'): item['Attr_Min_Custom_Height'] = txt.replace('Min. Custom Height:', '').strip()
            elif txt.startswith('O/A Height:'): item['Attr_OA_Height'] = txt.replace('O/A Height:', '').strip()
            elif txt.startswith('Chain Length:'): item['Attr_Chain_Length'] = txt.replace('Chain Length:', '').strip()
            elif txt.startswith('Shade Details:'): item['Attr_Shade_Details'] = txt.replace('Shade Details:', '').strip()
            elif txt.startswith('Base:'): item['Attr_Base'] = txt.replace('Base:', '').strip()
            elif txt.startswith('Extension:'): item['Attr_Extension'] = txt.replace('Extension:', '').strip()
            elif txt.startswith('Backplate:'): item['Attr_Blackplate'] = txt.replace('Backplate:', '').strip()
            elif txt.startswith('Rating:'): item['Attr_Rating'] = txt.replace('Rating:', '').strip()
            elif txt.startswith('CFM Average:'): item['Attr_CFM_Average'] = txt.replace('CFM Average:', '').strip()
            elif txt.startswith('CFM Low:'): item['Attr_CFM_Low'] = txt.replace('CFM Low:', '').strip()
            elif txt.startswith('CFM High:'): item['Attr_CFM_High'] = txt.replace('CFM High:', '').strip()
            elif txt.startswith('CFM Wattage Low:'): item['Attr_CFM_Wattage_Low'] = txt.replace('CFM Wattage Low:', '').strip()
            elif txt.startswith('CFM Wattage High:'): item['Attr_CFM_Wattage_High'] = txt.replace('CFM Wattage High:', '').strip()
    except: pass

    return item

# --- MAIN EXECUTION ---

def main():
    global STOP_EXECUTION
    driver = get_driver()
    all_details = []
    
    try:
        print(f"--- Starting Scraping Job ---")

        for main_cat, urls in target_map.items():
            if STOP_EXECUTION: break 
            print(f"\n=== Category: {main_cat} ===")

            for url_idx, url in enumerate(urls):
                if STOP_EXECUTION: break 
                print(f"-> Collecting Links from {url}")
                
                # Phase 1: Link Collection
                collected_links = []
                try:
                    driver.get(url)
                    if url_idx == 0: 
                        time.sleep(3)
                        check_popup(driver)
                    
                    page_num = 1
                    while True:
                        if STOP_EXECUTION: break 

                        try:
                            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "li.product-card")))
                        except TimeoutException:
                            break
                        
                        # Fast scroll for list page (no need for visual rendering here)
                        driver.execute_script("window.scrollBy(0, 600);")
                        time.sleep(1.5)

                        cards = driver.find_elements(By.CSS_SELECTOR, "li.product-card .name a")
                        for card in cards:
                            l = card.get_attribute('href')
                            if l and l not in collected_links:
                                collected_links.append(l)
                        
                        print(f"      Page {page_num}: {len(collected_links)} links found.")

                        try:
                            next_btn = driver.find_element(By.CSS_SELECTOR, "button.next.reset")
                            if "display: none" in next_btn.get_attribute("style"): break
                            driver.execute_script("arguments[0].scrollIntoView(true); arguments[0].click();", next_btn)
                            time.sleep(4)
                            page_num += 1
                        except: break 

                except Exception as e:
                    print(f"   [!] Link Error: {e}")

                # Phase 2: Extraction
                if STOP_EXECUTION: break

                total_products = len(collected_links)
                print(f"   >>> Starting Extraction for {total_products} products...")

                for i, p_url in enumerate(collected_links, 1):
                    if STOP_EXECUTION: 
                        print("[!] Stopping Extraction Loop...")
                        break

                    try:
                        data = extract_product_details(driver, p_url, main_cat)
                        if data:
                            all_details.append(data)
                            # LOGGING FORMAT: Scraping [Current]/[Total] of [Category] -> [Name] -> [URL]
                            p_name = data.get('Product Name', 'Unknown')
                            print(f"Scraping {i}/{total_products} of {main_cat} -> {p_name} -> {p_url}")

                        # Autosave
                        if len(all_details) % AUTOSAVE_INTERVAL == 0:
                            save_to_excel(all_details, EXCEL_FILENAME)

                    except Exception as e:
                        print(f"      [!] Error: {e}")

    except Exception as e:
        print(f"\n[ERROR] Unexpected crash: {e}")

    finally:
        print("\n[SYSTEM] Closing Driver and Saving Final Data...")
        try:
            driver.quit()
        except: 
            print("[!] Driver already closed.")
        
        save_to_excel(all_details, EXCEL_FILENAME)
        print("[SYSTEM] Done.")
        sys.exit()

if __name__ == "__main__":
    main()