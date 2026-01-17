import pandas as pd
import time
from urllib.parse import urljoin
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup

# ================= CONFIGURATION =================
BASE_URL = "https://sarreid.com/233/catalog/" 
OUTPUT_FILE = "sarreid_product_data_final.xlsx"
# =================================================

def init_driver():
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled")
    # options.add_argument("--headless") # Uncomment to run without window
    driver = webdriver.Chrome(options=options)
    return driver

def clean_text(text):
    if text:
        return text.strip().replace('\xa0', ' ')
    return "N/A"

# ================= PHASE 1: CRAWLING LOGIC =================

def handle_pagination(driver, wait):
    """ Attempts to click 'Show All' to load all products. """
    try:
        show_all_xpath = "//a[contains(@href, 'page=all')] | //a[contains(., 'Show All')]"
        show_all_btns = driver.find_elements(By.XPATH, show_all_xpath)
        
        if show_all_btns:
            btn = show_all_btns[0]
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
            time.sleep(1) 
            try:
                btn.click()
            except:
                driver.execute_script("arguments[0].click();", btn)
            
            try:
                wait.until(EC.url_contains("page=all"))
                time.sleep(3) 
            except: pass
    except: pass

def collect_all_product_urls(driver, wait, actions):
    """ Navigates menus and collects all product URLs. """
    all_urls = []
    unique_check = set()
    
    target_menus = ["Furniture", "Accessories"]
    main_categories = []

    print("--- Collecting Categories ---")
    for menu_name in target_menus:
        try:
            menu_xpath = f"//ul[contains(@class, 'sf-menu')]//a[strong[text()='{menu_name}']]"
            menu_element = wait.until(EC.visibility_of_element_located((By.XPATH, menu_xpath)))
            actions.move_to_element(menu_element).perform()
            time.sleep(2) 
            
            sub_menu_items = menu_element.find_elements(By.XPATH, "./following-sibling::ul//li/a")
            for item in sub_menu_items:
                cat_name = item.get_attribute("title")
                cat_url = item.get_attribute("href")
                if cat_url and "javascript" not in cat_url:
                    main_categories.append({
                        "Category Name": cat_name,
                        "Category URL": cat_url
                    })
        except Exception as e:
            print(f"Error menu {menu_name}: {e}")

    print(f"\nFound {len(main_categories)} categories. Starting URL collection...\n")

    for cat in main_categories:
        print(f"Processing Category: {cat['Category Name']}")
        driver.get(cat['Category URL'])
        
        sub_cat_urls = []
        try:
            sub_elements = driver.find_elements(By.CSS_SELECTOR, "div#subcat_categories_prd h3.name a")
            for sc in sub_elements:
                sc_url = sc.get_attribute("href")
                if sc_url and "index.php" in sc_url:
                    sub_cat_urls.append({"name": sc.text.strip(), "url": sc_url})
        except: pass

        def scrape_page(c_name):
            handle_pagination(driver, wait)
            try:
                links = driver.find_elements(By.CSS_SELECTOR, "h3.name a")
                for link in links:
                    p_url = link.get_attribute("href")
                    if p_url and "product_info.php" in p_url:
                         if p_url not in unique_check:
                            unique_check.add(p_url)
                            all_urls.append({"Category": c_name, "Product Url": p_url})
            except: pass

        scrape_page(cat['Category Name'])
        
        for sub in sub_cat_urls:
            print(f"   - Visiting Sub-Category: {sub['name']}")
            driver.get(sub['url'])
            scrape_page(f"{cat['Category Name']} > {sub['name']}")
                    
    return all_urls

# ================= PHASE 2: DETAIL EXTRACTION LOGIC (N/A Handling) =================

def extract_product_details(driver, category, product_url):
    driver.get(product_url)
    
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    data = {}
    
    # --- 1. PRE-FILL ALL FIELDS WITH "N/A" ---
    required_fields = [
        'Category', 'Product Url', 'Product Name', 'SKU', 'Price', 'Brand',
        'Short Description', 'Color / Style', 'Dimensions', 'Total Weight', 
        'Feature', 'Ships KD', 'On Hand', 'ETA', 'Full Description',
        'Image1', 'Image2', 'Image3', 'Image4'
    ]
    for field in required_fields:
        data[field] = "N/A"

    # --- 2. OVERWRITE WITH ACTUAL DATA IF FOUND ---
    
    # Basic Info
    data['Category'] = category
    data['Product Url'] = product_url

    # Product Name
    try:
        name_span = soup.find('span', itemprop='name')
        if name_span:
            data['Product Name'] = clean_text(name_span.get_text())
        else:
            url_link = soup.find('a', itemprop='url')
            if url_link:
                data['Product Name'] = clean_text(url_link.get_text())
            else:
                h1 = soup.find('h1')
                if h1:
                    h1_text = clean_text(h1.get_text())
                    if "Login" not in h1_text:
                        data['Product Name'] = h1_text.split('[')[0].strip()
    except: pass

    # SKU
    try:
        sku_inner = soup.find('span', itemprop='model')
        if sku_inner:
             data['SKU'] = clean_text(sku_inner.get_text())
        else:
            sku_outer = soup.find('span', class_='model')
            if sku_outer:
                data['SKU'] = clean_text(sku_outer.get_text()).replace('[', '').replace(']', '')
    except: pass

    # Price
    try:
        msrp_div = soup.find('div', class_='msrp')
        if msrp_div:
            data['Price'] = clean_text(msrp_div.get_text()).replace('MSRP :', '').strip()
    except: pass

    data['Brand'] = "Sarried Ltd"

    # Description Block (Specs)
    specs_div = None
    full_desc_div = None
    
    desc_divs = soup.find_all('div', class_='inner', itemprop='description')
    for d in desc_divs:
        if d.find('strong'):
            specs_div = d
        else:
            full_desc_div = d

    specs_map = {
        'Description': 'Short Description', 
        'Color / Style': 'Color / Style', 
        'Dimensions': 'Dimensions', 
        'Total Weight': 'Total Weight', 
        'Feature': 'Feature', 
        'Ships KD': 'Ships KD', 
        'On Hand': 'On Hand', 
        'ETA': 'ETA'
    }

    if specs_div:
        for key, col_name in specs_map.items():
            try:
                target_strong = specs_div.find('strong', string=lambda text: text and key in text)
                if target_strong:
                    next_node = target_strong.next_sibling
                    if next_node:
                        val = clean_text(str(next_node))
                        if val:
                            data[col_name] = val
            except: pass
    
    # Full Description
    if full_desc_div:
        val = clean_text(full_desc_div.get_text())
        if val:
            data['Full Description'] = val

    # --- IMAGE EXTRACTION FIX ---
    
    def clean_image_link(raw_href):
        """
        Input: images/../../../hiRes/53163_3.jpg
        Output: https://sarreid.com/hiRes/53163_3.jpg
        """
        if not raw_href: return None
        if "hiRes" in raw_href:
            # Split by 'hiRes' and take the last part (e.g., /53163_3.jpg)
            # Then reconstruct using the clean root domain
            path_part = raw_href.split("hiRes")[-1]
            return f"https://sarreid.com/hiRes{path_part}"
        return None

    # Image 1 (Main)
    try:
        img_container_1 = soup.find('div', class_='photoset-row cols-1')
        if img_container_1:
            a_tag = img_container_1.find('a')
            if a_tag and a_tag.get('href'):
                clean_url = clean_image_link(a_tag['href'])
                if clean_url:
                    data['Image1'] = clean_url
    except: pass

    # Images 2-4 (Gallery)
    try:
        gallery_containers = soup.find_all('div', class_='photoset-row cols-4')
        img_count = 2
        for container in gallery_containers:
            anchors = container.find_all('a')
            for a in anchors:
                if img_count > 4: break
                
                raw_href = a.get('href')
                if raw_href:
                    clean_url = clean_image_link(raw_href)
                    if clean_url:
                        data[f'Image{img_count}'] = clean_url
                        img_count += 1
                        
            if img_count > 4: break
    except: pass

    return data

# ================= MAIN EXECUTION =================

def main():
    driver = init_driver()
    wait = WebDriverWait(driver, 15)
    actions = ActionChains(driver)
    
    final_data = []

    try:
        # STEP 1: Crawl URLs
        driver.get(BASE_URL)
        print("=== PHASE 1: COLLECTING PRODUCT URLS ===")
        url_list = collect_all_product_urls(driver, wait, actions)
        print(f"\n--- Phase 1 Complete. Collected {len(url_list)} unique product URLs. ---")
        
        # STEP 2: Extract Details
        print("\n=== PHASE 2: EXTRACTING PRODUCT DETAILS ===")
        total = len(url_list)
        
        for index, item in enumerate(url_list):
            try:
                details = extract_product_details(driver, item['Category'], item['Product Url'])
                final_data.append(details)
                
                # Log progress
                p_name = details.get('Product Name', 'N/A')
                p_sku = details.get('SKU', 'N/A')
                print(f"[{index + 1}/{total}] {item['Category']} -> {p_name} -> {p_sku}")
                
                # Save every 50 products
                if (index + 1) % 50 == 0:
                    print(f"... Auto-saving backup at {index + 1} products ...")
                    pd.DataFrame(final_data).to_excel(OUTPUT_FILE, index=False)
                    
            except Exception as e:
                print(f"Error extracting {item['Product Url']}: {e}")

    except KeyboardInterrupt:
        print("\n!!! Interrupted by User (Ctrl+C). Saving data... !!!")
        
    except Exception as e:
        print(f"\n!!! Critical Script Error: {e} !!!")
        
    finally:
        driver.quit()
        if final_data:
            print(f"Saving final data to {OUTPUT_FILE}...")
            pd.DataFrame(final_data).to_excel(OUTPUT_FILE, index=False)
            print("Done.")
        else:
            print("No data was collected.")

if __name__ == "__main__":
    main()