import time
import signal
import sys
import pandas as pd
import traceback
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# =========================================
# CONFIG
# =========================================
BASE_URL = "https://tedboerner.com/furniture/all-furniture"
OUTPUT_FILE = "tedboerner_products.xlsx"
AUTOSAVE_INTERVAL = 50
CATEGORY_BREADCRUMB = "Home > Furniture > View All Furniture"

# =========================================
# SETUP SELENIUM
# =========================================
options = Options()
options.add_argument("--start-maximized")
options.add_argument("--disable-notifications")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 20)

# =========================================
# DATA HANDLING
# =========================================
data = []

def save_data():
    if not data:
        return
    df = pd.DataFrame(data)
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"💾 Progress saved ({len(data)} products) → {OUTPUT_FILE}")

def signal_handler(sig, frame):
    print("\n🛑 Ctrl+C detected! Saving progress before exit...")
    save_data()
    driver.quit()
    sys.exit(0)

signal.signal(signal.SIGINT, signal_handler)

# =========================================
# GET PRODUCT NAME (7 layouts)
# =========================================
def get_product_name(driver):
    selectors = [
        "h2.article_anywhere_title",
        "h5 > strong",
        "p > strong",
        "div > strong",
        "h2.title > span",
        "div",
    ]
    for sel in selectors:
        try:
            elem = driver.find_element(By.CSS_SELECTOR, sel)
            name = elem.text.strip()
            if name:
                return name
        except:
            continue
    return ""

# =========================================
# GET FULL DESCRIPTION (2 layouts)
# =========================================
def get_description(driver):
    # Layout 1
    try:
        article = driver.find_element(By.CSS_SELECTOR, "div.article_anywhere, div[itemprop='articleBody']")
        children = article.find_elements(By.XPATH, "./p | ./div")
        paragraphs = []
        for c in children:
            try:
                a_tags = c.find_elements(By.TAG_NAME, "a")
                if a_tags and any(".pdf" in a.get_attribute("href") for a in a_tags):
                    continue
                text = c.text.strip()
                if text:
                    paragraphs.append(text)
            except:
                continue
        if paragraphs:
            return "\n".join(paragraphs)
    except:
        pass
    
    # Layout 2
    try:
        div_blocks = driver.find_elements(By.CSS_SELECTOR, "div.article-anywhere, div[itemprop='articleBody'] div")
        paragraphs = [d.text.strip() for d in div_blocks if d.text.strip() and d.text.strip() != "\xa0"]
        if paragraphs:
            return "\n".join(paragraphs)
    except:
        pass

    return ""

# =========================================
# GET TAB DATA (Dimensions / Options / Materials)
# =========================================
def get_tabs(driver):
    tab_data = {"Dimensions": "", "Options/Choices": "", "Materials/Finishes": ""}
    keys = list(tab_data.keys())

    # Try layout 1: uk-tab
    try:
        tabs = driver.find_elements(By.CSS_SELECTOR, "ul.uk-tab li a")
        for i, tab in enumerate(tabs[:3]):
            driver.execute_script("arguments[0].click();", tab)
            time.sleep(1)
            try:
                html = driver.find_element(By.CSS_SELECTOR, "li.uk-active div.uk-panel").get_attribute("outerHTML")
                tab_data[keys[i]] = html
            except:
                continue
    except:
        pass

    # Try layout 2: uk-subnav-pill
    try:
        tabs = driver.find_elements(By.CSS_SELECTOR, "ul.uk-subnav.uk-subnav-pill li a")
        for i, tab in enumerate(tabs[:3]):
            driver.execute_script("arguments[0].click();", tab)
            time.sleep(1)
            try:
                html = driver.find_element(By.CSS_SELECTOR, "li.uk-active div.uk-panel").get_attribute("outerHTML")
                tab_data[keys[i]] = html
            except:
                continue
    except:
        pass

    return tab_data

# =========================================
# GET PRODUCT IMAGE
# =========================================
def get_product_image(driver):
    try:
        img_elem = driver.find_element(By.CSS_SELECTOR, "div.tp-bgimg.defaultimg")
        img = img_elem.get_attribute("data-lazyload")
        if not img or img.lower() == "undefined":
            img = img_elem.get_attribute("data-src")
        if not img or img.lower() == "undefined":
            img = img_elem.get_attribute("src")
        return img or ""
    except:
        return ""

# =========================================
# SCRAPE PRODUCT LIST PAGE
# =========================================
try:
    print("Opening product listing page...")
    driver.get(BASE_URL)
    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[data-grid-prepared='true']")))
    time.sleep(2)

    product_blocks = driver.find_elements(By.CSS_SELECTOR, "div[data-grid-prepared='true']")
    product_urls = [block.find_element(By.CSS_SELECTOR, "a.uk-position-cover").get_attribute("href") for block in product_blocks]
    print(f"✅ Found {len(product_urls)} products on the listing page.\n")

except Exception as e:
    print("❌ Error while collecting product URLs:", e)
    driver.quit()
    sys.exit(1)

# =========================================
# SCRAPE PRODUCT DETAILS
# =========================================
for idx, url in enumerate(product_urls, start=1):
    try:
        driver.get(url)
        time.sleep(2)

        name = get_product_name(driver)
        if not name:
            print(f"⚠️ Skipping product {idx}, no name found ({url})")
            continue

        print(f"🔹 Scraping ({idx}/{len(product_urls)}) → {name} | URL: {url}")

        full_desc = get_description(driver)

        # Tearsheet
        tear = ""
        try:
            tear = driver.find_element(By.XPATH, "//a[contains(@href, '.pdf')]").get_attribute("href")
        except:
            pass

        # Tabs
        tab_data = get_tabs(driver)

        # Image
        img = get_product_image(driver)

        product_info = {
            "Category": CATEGORY_BREADCRUMB,
            "Product URL": url,
            "Product Name": name,
            "SKU": "",
            "Brand": "Ted Boerner",
            "Description (Text)": full_desc,
            "Tearsheet": tear,
            "Dimensions (HTML)": tab_data["Dimensions"],
            "Options/Choices (HTML)": tab_data["Options/Choices"],
            "Materials/Finishes (HTML)": tab_data["Materials/Finishes"],
            "Image URL": img
        }

        data.append(product_info)

        if idx % AUTOSAVE_INTERVAL == 0:
            save_data()

    except Exception as e:
        print(f"⚠️ Error scraping product {idx} ({url}): {e}")
        traceback.print_exc()
        continue

# =========================================
# FINAL SAVE
# =========================================
save_data()
driver.quit()
print("✅ Scraping completed successfully.")
