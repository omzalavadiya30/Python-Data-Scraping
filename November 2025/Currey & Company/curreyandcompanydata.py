import time
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, ElementClickInterceptedException

# --- CONFIGURATION ---
CATEGORY_URLS = {
    "Lighting": [
        "https://www.curreyandcompany.com/c/lighting/",
        "https://www.curreyandcompany.com/c/lighting/chandeliers/",
        "https://www.curreyandcompany.com/c/lighting/pendants/",
        "https://www.curreyandcompany.com/c/lighting/multi-drop-pendants/",
        "https://www.curreyandcompany.com/c/lighting/lanterns/",
        "https://www.curreyandcompany.com/c/lighting/semi-flush-mounts/",
        "https://www.curreyandcompany.com/c/lighting/flush-mounts/",
        "https://www.curreyandcompany.com/c/lighting/bath-bars/",
        "https://www.curreyandcompany.com/c/lighting/wall-sconces/",
        "https://www.curreyandcompany.com/c/lighting/table-lamps/",
        "https://www.curreyandcompany.com/c/lighting/floor-lamps/",
        "https://www.curreyandcompany.com/c/lighting/post-lights/",
        "https://www.curreyandcompany.com/c/new/lighting/",
        "https://www.curreyandcompany.com/c/lighting/bath-collection/",
        "https://www.curreyandcompany.com/c/lighting/cordless-lighting/",
        "https://www.curreyandcompany.com/c/bestsellers/lighting/",
        "https://www.curreyandcompany.com/c/sale/lighting/",
        "https://www.curreyandcompany.com/c/lighting-accessories/additional-chains-rods/",
        "https://www.curreyandcompany.com/c/lighting-accessories/cordless-lighting-bases-bulbs/",
        "https://www.curreyandcompany.com/c/lighting-accessories/chandelier-shades/",
        "https://www.curreyandcompany.com/c/lighting-accessories/lamp-shades/",
        "https://www.curreyandcompany.com/c/lighting-accessories/light-bulbs/"
    ],
    "Furniture": [
        "https://www.curreyandcompany.com/c/furniture/",
        "https://www.curreyandcompany.com/c/furniture/chests-nightstands/",
        "https://www.curreyandcompany.com/c/furniture/cabinets-credenzas/",
        "https://www.curreyandcompany.com/c/furniture/desks-vanities/",
        "https://www.curreyandcompany.com/c/furniture/tables/",
        "https://www.curreyandcompany.com/c/furniture/accent-pieces/",
        "https://www.curreyandcompany.com/c/furniture/ottomans-benches/",
        "https://www.curreyandcompany.com/c/furniture/accent-chairs/",
        "https://www.curreyandcompany.com/c/furniture/dining-chairs/",
        "https://www.curreyandcompany.com/c/furniture/bar-counter-stools/",
        "https://www.curreyandcompany.com/c/furniture/bath-vanities/",
        "https://www.curreyandcompany.com/c/new/furniture/",
        "https://www.curreyandcompany.com/c/bestsellers/furniture/",
        "https://www.curreyandcompany.com/c/sale/furniture/",
        "https://www.curreyandcompany.com/c/furniture-accessories/glass-tops/",
        "https://www.curreyandcompany.com/c/furniture-accessories/bath-vanity-sidesplashes/",
        "https://www.curreyandcompany.com/c/furniture-accessories/hardware/"
    ],
    "Accessories": [
        "https://www.curreyandcompany.com/c/accessories/",
        "https://www.curreyandcompany.com/c/accessories/mirrors/",
        "https://www.curreyandcompany.com/c/accessories/boxes-trays/",
        "https://www.curreyandcompany.com/c/accessories/objects-sculptures/",
        "https://www.curreyandcompany.com/c/accessories/vases-jars-bowls/",
        "https://www.curreyandcompany.com/c/new/accessories/",
        "https://www.curreyandcompany.com/c/bestsellers/accessories/",
        "https://www.curreyandcompany.com/c/sale/accessories/"
    ],
    "Outdoor": [
        "https://www.curreyandcompany.com/c/outdoor/",
        "https://www.curreyandcompany.com/c/outdoor/tables/",
        "https://www.curreyandcompany.com/c/outdoor/seating/",
        "https://www.curreyandcompany.com/c/outdoor/accessories/",
        "https://www.curreyandcompany.com/c/outdoor/lighting/",
        "https://www.curreyandcompany.com/c/outdoor/planters/",
        "https://www.curreyandcompany.com/c/new/outdoor/",
        "https://www.curreyandcompany.com/c/bestsellers/outdoor/",
        "https://www.curreyandcompany.com/c/sale/outdoor/"
    ],
    "New": [
        "https://www.curreyandcompany.com/c/new/",
        "https://www.curreyandcompany.com/c/new-introductions/fall-2025/",
        "https://www.curreyandcompany.com/c/market-bestsellers/"
    ],
    "One of a Kind": [
        "https://www.curreyandcompany.com/c/one-of-a-kind/",
        "https://www.curreyandcompany.com/c/one-of-a-kind/vessels/",
        "https://www.curreyandcompany.com/c/one-of-a-kind/objets/"
    ],
    "Collaborations": [
        "https://www.curreyandcompany.com/collaborations/aviva-stanoff-collection/",
        "https://www.curreyandcompany.com/collaborations/barry-goralnick-collection/",
        "https://www.curreyandcompany.com/collaborations/bunny-williams-collection/",
        "https://www.curreyandcompany.com/collaborations/hiroshi-koshitaka-collection/",
        "https://www.curreyandcompany.com/collaborations/jamie-beckwith-collection/",
        "https://www.curreyandcompany.com/collaborations/lillian-august-collection/",
        "https://www.curreyandcompany.com/collaborations/marjorie-skouras-collection/",
        "https://www.curreyandcompany.com/collaborations/sasha-bikoff-collection/",
        "https://www.curreyandcompany.com/collaborations/suzanne-duin-collection/",
        "https://www.curreyandcompany.com/collaborations/winterthur-collection/"
    ],
    "Inspiration": [
        "https://www.curreyandcompany.com/c/explore-our-booth/",
        "https://www.curreyandcompany.com/c/collections/amber-green/",
        "https://www.curreyandcompany.com/c/collections/collection-no.-5/",
        "https://www.curreyandcompany.com/c/collections/critter-corner/",
        "https://www.curreyandcompany.com/c/collections/in-the-wild/",
        "https://www.curreyandcompany.com/c/collections/little-luxuries/",
        "https://www.curreyandcompany.com/c/collections/the-neutral-edit/",
        "https://www.curreyandcompany.com/c/collections/perfectly-wicked/",
        "https://www.curreyandcompany.com/c/collections/the-polo-club/",
        "https://www.curreyandcompany.com/c/collections/st.-barts/",
        "https://www.curreyandcompany.com/c/collections/sustainability/"
    ],
    "Sale": [
        "https://www.curreyandcompany.com/c/sale/"
    ]
}

OUTPUT_FILE = "currey_products_detailed.xlsx"

def setup_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    # options.add_argument("--headless") # Uncomment to run in background
    driver = webdriver.Chrome(options=options)
    return driver

# --- PHASE 1: URL COLLECTION FUNCTIONS ---

def scrape_products_on_page(driver):
    """ Extracts product URLs from the current view. """
    product_urls = set()
    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'a[data-title="Card-CTA"]'))
        )
        products = driver.find_elements(By.CSS_SELECTOR, 'a[data-title="Card-CTA"]')
        for p in products:
            url = p.get_attribute('href')
            if url:
                product_urls.add(url)
    except TimeoutException:
        pass
    return list(product_urls)

def handle_pagination(driver):
    """ Checks for pagination, verifies if 'Next' is enabled, and clicks it. """
    next_button = None
    
    # 1. Search for the "Next" button (Try both styles)
    try:
        # Collaboration Style (Slider)
        btn = driver.find_element(By.XPATH, '//button[contains(@class, "collection-nav-button") and @data-text="Forward"]')
        if btn.is_displayed(): next_button = btn
    except NoSuchElementException:
        pass

    if not next_button:
        try:
            # Standard Style (Grid)
            btn = driver.find_element(By.XPATH, '//button[contains(@class, "navButton") and @data-text="Forward"]')
            if btn.is_displayed(): next_button = btn
        except NoSuchElementException:
            pass

    if not next_button: return False
    
    # Check if disabled
    if not next_button.is_enabled(): return False
    button_class = next_button.get_attribute("class")
    if "disabled" in button_class.lower() or "Mui-disabled" in button_class: return False

    # Click
    try:
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", next_button)
        time.sleep(1)
        try:
            next_button.click()
        except ElementClickInterceptedException:
            driver.execute_script("arguments[0].click();", next_button)
        print("    -> Clicking Next Page...")
        time.sleep(4)
        return True
    except Exception as e:
        print(f"    ! Pagination Error: {e}")
        return False

# --- PHASE 2: DETAIL EXTRACTION FUNCTIONS ---

def safe_extract(driver, by, value, attribute="text"):
    """ Helper to extract text or attribute safely. """
    try:
        element = driver.find_element(by, value)
        if attribute == "text":
            return element.text.strip()
        elif attribute == "outerHTML":
            return element.get_attribute("outerHTML")
        else:
            return element.get_attribute(attribute)
    except NoSuchElementException:
        return "N/A"

def get_spec_value_from_row(driver, label_name):
    """ 
    Robustly finds a specification value based on the Label Name.
    Uses the 'paragraph-2b-reg' class structure found in the HTML.
    """
    try:
        # Logic: Find the label text inside a span, ensure it's inside a row container, 
        # then get the value from the sibling div.
        # We use normalize-space() to ignore extra whitespace/newlines.
        xpath = f"//div[contains(@class, 'paragraph-2b-reg')]//span[normalize-space(text())='{label_name}']/ancestor::div[contains(@class, 'justify-between')]/div[2]/span"
        
        element = driver.find_element(By.XPATH, xpath)
        return element.text.strip()
    except NoSuchElementException:
        return "N/A"

def ensure_accordion_visible(driver, text_identifier):
    """
    Scrolls to an accordion header. 
    Since they are open by default, we just need to ensure they are in view.
    """
    try:
        # Matches "Dimensions", "Lighting Specifications", "Furniture Specifications", etc.
        xpath = f"//span[contains(@class, 'paragraph-1a-lg') and contains(text(), '{text_identifier}')]"
        header = driver.find_element(By.XPATH, xpath)
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", header)
        time.sleep(0.5)
        return True
    except NoSuchElementException:
        return False

def scrape_single_product(driver, url):
    """ Visits a product page and extracts all details. """
    driver.get(url)
    time.sleep(2) 

    data = {}

    # 1. Category Path & Product Name
    try:
        breadcrumbs = driver.find_element(By.CSS_SELECTOR, "div.breadcrumbs-wrapper").text.replace("\n", " > ")
        data['Category'] = breadcrumbs
    except:
        data['Category'] = "N/A"

    data['Product URL'] = url
    # Basic Info
    data['Product Name'] = safe_extract(driver, By.CSS_SELECTOR, 'h4.headline-1c')
    data['SKU'] = safe_extract(driver, By.CSS_SELECTOR, 'span.paragraph-2b-reg')
    data['Brand'] = "Currey & Company"
    data['Description'] = safe_extract(driver, By.CSS_SELECTOR, 'div.account-paragraph-s')


    # 2. Full HTML (Scroll needed)
    try:
        desc_container = driver.find_element(By.ID, "additional")
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", desc_container)
        time.sleep(1)
        data['Full Description HTML'] = desc_container.get_attribute('outerHTML')
    except NoSuchElementException:
        data['Full Description HTML'] = "N/A"

    # 3. EXTRACT DIMENSIONS
    # Scroll to Dimensions section
    ensure_accordion_visible(driver, "Dimensions")
    
    dim_fields = [
        "Overall", "Item Weight", "Lamp Base", "Lamp Body", 
        "Cord", "Shade Top", "Shade Bottom", "Shade Height"
    ]
    for field in dim_fields:
        data[field] = get_spec_value_from_row(driver, field)

    # 4. EXTRACT SPECIFICATIONS
    # Scroll to Specifications (Dynamic: "Lighting Specifications", "Furniture Specifications", etc.)
    # We look for *any* header containing "Specifications"
    ensure_accordion_visible(driver, "Specifications")

    spec_fields = [
        "Finish", "Material", "Hardware Details", "Floor Protection", 
        "Light Source", "Light Direction", "Voltage", "Fixture Type", "Finial"
    ]
    for field in spec_fields:
        data[field] = get_spec_value_from_row(driver, field)

    # 5. Tearsheet
    try:
        tear_link = driver.find_element(By.CSS_SELECTOR, 'a[data-text="Tearsheet"]')
        data['Tearsheet'] = tear_link.get_attribute('href')
    except NoSuchElementException:
        data['Tearsheet'] = "N/A"

    # 6. Images
    try:
        img1 = driver.find_element(By.CSS_SELECTOR, 'div.image-gallery-center img')
        data['Image 1'] = img1.get_attribute('src')
    except NoSuchElementException:
        data['Image 1'] = "N/A"

    try:
        thumbnails = driver.find_elements(By.CSS_SELECTOR, 'div.image-gallery-thumbnails button img')
        img_count = 2
        for i in range(1, len(thumbnails)): 
            if img_count > 4: break
            src = thumbnails[i].get_attribute('src')
            if "width=100" in src:
                src = src.replace("width=100", "width=1600")
            data[f'Image {img_count}'] = src
            img_count += 1
    except:
        pass
    
    for j in range(img_count, 5):
        data[f'Image {j}'] = "N/A"

    return data

def save_data(data_list, filename):
    if not data_list: return
    df = pd.DataFrame(data_list)
    if os.path.exists(filename):
        existing_df = pd.read_excel(filename)
        combined_df = pd.concat([existing_df, df], ignore_index=True)
        combined_df.drop_duplicates(subset=["Product URL"], keep='last', inplace=True)
        combined_df.to_excel(filename, index=False)
    else:
        df.to_excel(filename, index=False)
    print(f"   >>> Auto-saved {len(df)} rows to {filename}")

# --- MAIN CONTROLLER ---

def main():
    driver = setup_driver()
    final_product_list = [] 
    
    print("Starting Detailed Scraper V2. Press Ctrl+C to stop and save safely.\n")

    try:
        for main_category, urls in CATEGORY_URLS.items():
            print(f"### PROCESSING CATEGORY: {main_category} ###")
            
            # --- STEP 1: Collect URLs ---
            category_product_urls = set()
            
            for url in urls:
                print(f"   > Collecting URLs from: {url}")
                driver.get(url)
                time.sleep(3)
                
                page_num = 1
                while True:
                    products = scrape_products_on_page(driver)
                    
                    for p in products:
                        category_product_urls.add(p)
                    
                    print(f"     Page {page_num}: Found {len(products)} products. (Total Unique so far: {len(category_product_urls)})")
                    
                    if not products and page_num > 1: break
                    if not handle_pagination(driver): break
                    page_num += 1
            
            category_product_urls = list(category_product_urls)
            total_products = len(category_product_urls)
            print(f"\n   -> Total products found in {main_category}: {total_products}")
            print("   -> Starting Detail Extraction...\n")

            # --- STEP 2: Extract Details ---
            for index, p_url in enumerate(category_product_urls):
                try:
                    p_data = scrape_single_product(driver, p_url)
                    
                    final_product_list.append(p_data)
                    
                    print(f"Scraping {index + 1}/{total_products} -> {p_data.get('Product Name', 'Unknown')}")

                    # Auto-save every 50 products
                    if len(final_product_list) % 50 == 0:
                        save_data(final_product_list, OUTPUT_FILE)
                        final_product_list = [] 

                except Exception as e:
                    print(f"   ! Error on {p_url}: {e}")
                    continue

    except KeyboardInterrupt:
        print("\n\n!!! Interrupted by User (Ctrl+C) !!!")
        print("Saving collected data before exiting...")
    
    except Exception as e:
        print(f"\n\n!!! Critical Error: {e} !!!")

    finally:
        if final_product_list:
            save_data(final_product_list, OUTPUT_FILE)
        
        print("\n--- Scraping Finished. Closing Driver. ---")
        driver.quit()

if __name__ == "__main__":
    main()