import os
import time
import sys
import signal
import psutil
import pandas as pd
from urllib.parse import unquote

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.action_chains import ActionChains

from webdriver_manager.chrome import ChromeDriverManager

# --------------------------------------------------
# CONFIG
# --------------------------------------------------
BASE_URL = 'https://gabby.com/'
OUTPUT_FILE = 'gabby_products_details.xlsx'

SAVE_CHECKPOINT = 50
MAX_RETRIES = 3

CHECKPOINT_FILE = "gabby_checkpoint.txt"
RESTART_MEMORY_MB = 1200   # restart browser if Chrome > 1.2 GB RAM

# --------------------------------------------------
# SELECTORS (UNCHANGED FROM YOUR SCRIPT)
# --------------------------------------------------
MAIN_MENU_TRIGGER_SELECTOR = 'a[menu-trigger]'
SUB_CATEGORY_SELECTOR = 'a[href^="/collections/"]'
PAGINATION_NEXT_SELECTOR = 'a[aria-label="Next page"]'
PRODUCT_LISTING_SELECTOR = '#product-grid h3 a'

BREADCRUMB_SELECTOR = 'ol.flex.items-center'
PRODUCT_NAME_SELECTOR = 'h1.h4.text-t-brand-foreground-secondary'
SKU_SELECTOR = 'p[id^="Sku-template--"]'
DESCRIPTION_SELECTOR = 'div.inline-richtext > p'
FULL_DESCRIPTION_HTML_SELECTOR = 'section[id^="ProductInfo-template--"]'

MORE_INFO_ACCORDION_LABEL = 'label[for*="collapsible_tab_Vm63mW"]'
MORE_INFO_ACCORDION_CONTENT = 'div[id*="collapsible_tab_Vm63mW"] div[class*="pb-5"]'

WARRANTY_ACCORDION_LABEL = 'label[for*="collapsible-row-1"]'
WARRANTY_ACCORDION_CONTENT = 'div[id*="collapsible-row-1"] div[class*="pb-5"]'

FEATURES_ACCORDION_LABEL = 'label[for*="dimensions_tab_FMpAdN"]'
FEATURES_MODAL_BUTTON = 'button[aria-controls="product-specs-modal"]'
FEATURES_MODAL_SELECTOR = 'modal-dialog[id="product-specs-modal"]'
FEATURES_MODAL_SCROLL_AREA = 'modal-dialog[id="product-specs-modal"] div.overflow-y-auto'
FEATURES_MODAL_SECTION_SELECTOR = 'div.specs-attributes-section'
FEATURES_MODAL_ROWS_SELECTOR = 'div.grid.grid-cols-2.items-center > span'
FEATURES_MODAL_HTML_SELECTOR = 'modal-dialog[id="product-specs-modal"] div[class^="grid grid-cols-1 gap-y-lg"]'
FEATURES_MODAL_CLOSE_BUTTON = 'modal-dialog[id="product-specs-modal"] svg-icon[src="icon-close"]'

ACCORDION_FEATURES_CONTENT_SELECTOR = 'div[id*="dimensions_tab_FMpAdN"]'
IMAGE_THUMBNAIL_SELECTOR = 'swiper-container[id*="-thumbs-swiper"] img'

# --------------------------------------------------

all_product_data = []
driver_instance = None

# --------------------------------------------------
# CHECKPOINT
# --------------------------------------------------
def save_checkpoint(idx):
    with open(CHECKPOINT_FILE,"w") as f:
        f.write(str(idx))

def load_checkpoint():
    try:
        return int(open(CHECKPOINT_FILE).read())
    except:
        return 0

# --------------------------------------------------
# HELPERS
# --------------------------------------------------
def decode_url_if_encoded(url):
    return unquote(url) if url and ("%3" in url or "%2" in url) else url

def save_data():
    df = pd.DataFrame(all_product_data)
    df.drop_duplicates(subset=["Product URL"], inplace=True)
    df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")
    print(f"\n✅ Saved {len(df)} products")

def safe_get_text(driver, selector):
    try:
        return driver.find_element(By.CSS_SELECTOR, selector).text.strip()
    except:
        return None

def safe_get_html(driver, selector):
    try:
        el = driver.find_element(By.CSS_SELECTOR, selector)
        return el.get_attribute('innerHTML')
    except:
        return None

def safe_get_element_text(el):
    try:
        return el.text.strip()
    except:
        return ""

# --------------------------------------------------
# MEMORY WATCHDOG
# --------------------------------------------------
def should_restart_browser():
    chrome = [
        p for p in psutil.process_iter(['name','memory_info'])
        if p.info['name'] and 'chrome' in p.info['name'].lower()
    ]
    total = sum(
        p.info['memory_info'].rss for p in chrome
    ) / (1024*1024)

    print(f"🧠 Chrome RAM: {int(total)} MB")

    return total > RESTART_MEMORY_MB

# --------------------------------------------------
# DRIVER
# --------------------------------------------------
def create_driver():
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless=new')
    options.add_argument('--disable-gpu')
    options.add_argument('--window-size=1920,1080')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.implicitly_wait(5)
    return driver

def restart_driver(old):
    try:
        old.quit()
    except:
        pass
    time.sleep(2)
    return create_driver()

# --------------------------------------------------
# CATEGORY SCRAPE
# --------------------------------------------------
def get_category_links(driver):
    driver.get(BASE_URL)
    time.sleep(2)

    actions = ActionChains(driver)
    categories = {}

    triggers = driver.find_elements(By.CSS_SELECTOR, MAIN_MENU_TRIGGER_SELECTOR)

    for trigger in triggers:
        try:
            actions.move_to_element(trigger).perform()
            time.sleep(1)

            dropdown_id = trigger.get_attribute('aria-controls')
            if not dropdown_id:
                continue

            subs = driver.find_elements(By.CSS_SELECTOR, f"#{dropdown_id} {SUB_CATEGORY_SELECTOR}")

            for elem in subs:
                url = elem.get_attribute('href')
                name = elem.text.strip()
                if url and name:
                    categories[url] = name
        except:
            continue

    return categories

def scrape_category_page(driver, url):
    urls = []

    while True:
        driver.get(url)

        WebDriverWait(driver,10).until(
            EC.presence_of_element_located((By.ID,"product-grid"))
        )

        page_urls = driver.execute_script(
            "return [...document.querySelectorAll('#product-grid h3 a')].map(e=>e.href);"
        )
        urls += page_urls

        try:
            nxt = driver.find_element(By.CSS_SELECTOR, PAGINATION_NEXT_SELECTOR)
            url = nxt.get_attribute('href')
            if not url:
                break
        except:
            break

    return urls

# --------------------------------------------------
# PRODUCT PAGE SCRAPE (YOUR ORIGINAL COLUMNS KEPT)
# --------------------------------------------------
def get_product_details(driver, product_url, category_name):

    driver.get(product_url)

    data = {
        'Category': '',
        'Product URL': product_url,
        'Product Name': '',
        'SKU': '',
        'Brand': 'Gabby',

        'Attr_Product Depth': '',
        'Attr_Product Height': '',
        'Attr_Product Width': '',
        'Attr_Product Weight': '',
        'Attr_Seat Depth': '',
        'Attr_Seat Height': '',
        'Attr_Seat Width': '',

        'Material': '',
        'Finish Family': '',
        'Collection Name': '',
        'Number of Shelves': '',

        'Description': '',
        'Full Description HTML': '',
        'More Information': '',
        'Warranty': '',
        'Details & Specifications HTML': '',

        'Image1': None,
        'Image2': None,
        'Image3': None,
        'Image4': None
    }

    # Breadcrumb
    try:
        bc = driver.find_elements(By.CSS_SELECTOR, f"{BREADCRUMB_SELECTOR} li")
        data['Category'] = " > ".join(e.text.strip() for e in bc if e.text.strip())
    except:
        pass

    data['Product Name'] = safe_get_text(driver, PRODUCT_NAME_SELECTOR)

    sku = safe_get_text(driver, SKU_SELECTOR)
    data['SKU'] = sku.replace("SKU:","").strip() if sku else None

    data['Description'] = safe_get_text(driver, DESCRIPTION_SELECTOR)
    data['Full Description HTML'] = safe_get_html(driver, FULL_DESCRIPTION_HTML_SELECTOR)

    # More Info
    try:
        driver.find_element(By.CSS_SELECTOR, MORE_INFO_ACCORDION_LABEL).click()
        time.sleep(0.3)
        data['More Information'] = safe_get_text(driver, MORE_INFO_ACCORDION_CONTENT)
    except:
        pass

    # Warranty
    try:
        driver.find_element(By.CSS_SELECTOR, WARRANTY_ACCORDION_LABEL).click()
        time.sleep(0.3)
        data['Warranty'] = safe_get_text(driver, WARRANTY_ACCORDION_CONTENT)
    except:
        pass

    # Specs
    try:
        driver.find_element(By.CSS_SELECTOR, FEATURES_ACCORDION_LABEL).click()
        time.sleep(0.4)

        specs = {}

        try:
            modal_btn = WebDriverWait(driver,3).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, FEATURES_MODAL_BUTTON))
            )
            modal_btn.click()

            modal = WebDriverWait(driver,5).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, FEATURES_MODAL_SELECTOR))
            )

            data['Details & Specifications HTML'] = safe_get_html(driver, FEATURES_MODAL_HTML_SELECTOR)

            sections = modal.find_elements(By.CSS_SELECTOR, FEATURES_MODAL_SECTION_SELECTOR)

        except:
            data['Details & Specifications HTML'] = safe_get_html(
                driver, ACCORDION_FEATURES_CONTENT_SELECTOR
            )
            sections = driver.find_elements(By.CSS_SELECTOR, FEATURES_MODAL_SECTION_SELECTOR)

        for sec in sections:
            rows = sec.find_elements(By.CSS_SELECTOR, FEATURES_MODAL_ROWS_SELECTOR)
            for i in range(0, len(rows), 2):
                key = safe_get_element_text(rows[i])
                val = safe_get_element_text(rows[i+1])
                if key:
                    specs[key.lower()] = val

        data['Attr_Product Depth'] = specs.get('product depth')
        data['Attr_Product Height'] = specs.get('product height')
        data['Attr_Product Width'] = specs.get('product width')
        data['Attr_Product Weight'] = specs.get('product weight')
        data['Attr_Seat Depth'] = specs.get('seat depth', specs.get('seat cushion depth'))
        data['Attr_Seat Height'] = specs.get('seat height')
        data['Attr_Seat Width'] = specs.get('seat width')

        data['Material'] = specs.get('material')
        data['Finish Family'] = specs.get('finish family')
        data['Collection Name'] = specs.get('collection name')
        data['Number of Shelves'] = specs.get('number of shelves')

        try:
            driver.find_element(By.CSS_SELECTOR, FEATURES_MODAL_CLOSE_BUTTON).click()
        except:
            pass

    except:
        pass

    # Images
    try:
        imgs = driver.find_elements(By.CSS_SELECTOR, IMAGE_THUMBNAIL_SELECTOR)
        urls = [decode_url_if_encoded(i.get_attribute("src")) for i in imgs if i.get_attribute("src")]

        for i in range(4):
            data[f'Image{i+1}'] = urls[i] if i < len(urls) else None

    except:
        pass

    return data

# --------------------------------------------------
# CTRL+C HANDLER
# --------------------------------------------------
def signal_handler(sig, frame):
    print("\nCTRL+C detected — saving data")
    save_data()
    if driver_instance:
        driver_instance.quit()
    sys.exit()

signal.signal(signal.SIGINT, signal_handler)

# --------------------------------------------------
# MAIN
# --------------------------------------------------
def main():

    global driver_instance

    driver = create_driver()
    driver_instance = driver

    categories = get_category_links(driver)

    all_links = []
    for url,name in categories.items():
        for p in scrape_category_page(driver, url):
            all_links.append((p, name))

    unique_products = list(dict(all_links).items())
    total = len(unique_products)

    start = load_checkpoint()
    print(f"\n🔄 Resuming from {start+1}/{total}\n")

    for idx,(url,cat) in enumerate(unique_products[start:], start=start):

        if should_restart_browser():
            driver = restart_driver(driver)
            driver_instance = driver

        for attempt in range(MAX_RETRIES):

            try:
                data = get_product_details(driver, url, cat)
                all_product_data.append(data)

                print(f"{idx+1}/{total} ✅ {data['Product Name']}")
                save_checkpoint(idx+1)
                break

            except Exception as e:
                print(f"{idx+1}/{total} ❌ attempt {attempt+1} — {e}")

                if attempt == MAX_RETRIES - 1:
                    all_product_data.append({
                        "Product URL": url,
                        "Product Name": "SCRAPE_FAILED"
                    })
                else:
                    driver = restart_driver(driver)
                    driver_instance = driver

        if (idx+1) % SAVE_CHECKPOINT == 0:
            save_data()

    save_data()

    try:
        os.remove(CHECKPOINT_FILE)
    except:
        pass

    driver.quit()

    print("\n🎉 SCRAPE COMPLETED")

# --------------------------------------------------
if __name__ == "__main__":
    main()
