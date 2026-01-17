import os
import time
import signal
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ================= CONFIG =================
BASE_URL = "https://domiziani.com"
START_URL = "https://domiziani.com/en"
OUTPUT_FILE = "domiziani_final_products.xlsx"
SAVE_INTERVAL = 50  # auto-save every 50 products

# ================= SELENIUM SETUP =================
options = Options()
# options.add_argument("--headless=new")  # uncomment for headless mode
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-gpu")
options.add_argument("--start-maximized")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 30)
driver.set_page_load_timeout(180)

data = []
stop_requested = False

# ================= SAFE EXIT HANDLER =================
def handle_exit(sig, frame):
    global stop_requested
    print("\n🛑 Ctrl+C detected! Saving progress before exit...")
    stop_requested = True
    save_data()
    driver.quit()
    raise SystemExit("✅ Exited safely after saving progress.")

signal.signal(signal.SIGINT, handle_exit)

def save_data():
    """Save collected data to Excel safely."""
    if not data:
        return
    df = pd.DataFrame(data)
    df.drop_duplicates(inplace=True)
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"💾 Progress saved! ({len(df)} total records)")

# ================= STEP 1: OPEN SITE =================
for attempt in range(3):
    try:
        print(f"🌐 Loading site (attempt {attempt+1})...")
        driver.get(START_URL)
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "a")))
        break
    except Exception as e:
        print(f"⚠️ Failed attempt {attempt+1}: {e}")
        if attempt == 2:
            driver.quit()
            raise SystemExit("❌ Could not load homepage after 3 attempts.")

# ================= STEP 2: COLLECT CATEGORY LINKS =================
print("🔍 Collecting category links...")
category_links = set()
for link in driver.find_elements(By.TAG_NAME, "a"):
    href = link.get_attribute("href")
    if href and href.startswith(BASE_URL + "/en/") and ("/furniture/" in href or "/patterns/" in href):
        if href.count("/") <= 6:
            category_links.add(href)

category_links = list(category_links)
print(f"✅ Found {len(category_links)} categories")

# ================= STEP 3: SCRAPE CATEGORY PRODUCTS =================
def scrape_category_products(category_url):
    """Get all product URLs and names from a category."""
    for attempt in range(3):
        try:
            driver.get(category_url)
            wait.until(EC.presence_of_element_located(
                (By.CSS_SELECTOR, "div.mansory-grid-image.v-2.m-t-1.crazy-load")
            ))

            cards = driver.find_elements(By.CSS_SELECTOR, "a.tall.img-gallery-cat")
            products = []
            for card in cards:
                href = card.get_attribute("href")
                name_el = card.find_elements(By.CSS_SELECTOR, "p.name-product-slide.title-cat")
                name = name_el[0].text.strip() if name_el else ""
                if href:
                    full_url = href if href.startswith("http") else BASE_URL + href
                    products.append((name, full_url))
            return list(set(products))
        except Exception as e:
            print(f"⚠️ Attempt {attempt+1} failed for {category_url}: {e}")
            time.sleep(5)
    print("❌ Skipping category after 3 failed attempts.")
    return []

# ================= STEP 4: SCRAPE PRODUCT DETAILS =================
def scrape_product_details(product_url):
    """Extract product details with category (mobile view) and other fields (desktop view)."""
    details = {
        "Category": "",
        "Product URL": product_url,
        "Product Name": "",
        "SKU": "",
        "Brand": "DOMIZIANI",
        "Description": "",
        "Datasheet": "",
        "Image1": "",
        "Image2": "",
        "Image3": "",
        "Image4": ""
    }

    try:
        # --- Load in desktop view first ---
        driver.set_window_size(1920, 1080)
        driver.get(product_url)
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        time.sleep(2)

        # --- Product name ---
        try:
            name_el = driver.find_element(By.CSS_SELECTOR, "h1.h2-c")
            details["Product Name"] = name_el.text.strip()
        except:
            pass

        # --- SKU ---
        try:
            sku_el = driver.find_element(By.CSS_SELECTOR, "p.cod_p")
            details["SKU"] = sku_el.text.strip()
        except:
            pass

        # --- Description ---
        try:
            desc_el = driver.find_element(By.CSS_SELECTOR, "div.desc_product")
            details["Description"] = desc_el.text.strip()
        except:
            pass

        # --- Datasheet ---
        try:
            datasheet_el = driver.find_element(By.CSS_SELECTOR, "div.list-links.v_2 a")
            href = datasheet_el.get_attribute("href")
            if href and href.startswith("/"):
                href = BASE_URL + href
            details["Datasheet"] = href
        except:
            pass

        # --- Image 1 (from box-primary) ---
        try:
            img1 = driver.find_element(By.CSS_SELECTOR, "div.box-primary > img.img_primary").get_attribute("src")
            if img1 and img1.startswith("/"):
                img1 = BASE_URL + img1
            details["Image1"] = img1
        except:
            pass

        # --- Image 2–4 (from mansory-grid-image) ---
        try:
            imgs = driver.find_elements(By.CSS_SELECTOR, "div.mansory-grid-image.p_product img.img-gallery")
            for i, img in enumerate(imgs[:3], start=2):
                src = img.get_attribute("src")
                if src and src.startswith("/"):
                    src = BASE_URL + src
                details[f"Image{i}"] = src
        except:
            pass

        # --- Now switch to mobile view to get category ---
        try:
            driver.set_window_size(400, 800)
            time.sleep(2)
            crumbs = driver.find_elements(By.CSS_SELECTOR, "div.breadcrumbs-page.mobile ul.breadcrumbs li span")
            if crumbs:
                details["Category"] = " | ".join([c.text.strip() for c in crumbs if c.text.strip()])
        except:
            pass

        return details

    except Exception as e:
        print(f"⚠️ Error scraping product: {e}")
        return details

# ================= STEP 5: MAIN SCRAPE LOOP =================
total_scraped = 0

for c_idx, cat_url in enumerate(category_links, start=1):
    if stop_requested:
        break

    print(f"\n🗂️ [{c_idx}/{len(category_links)}] Category: {cat_url}")
    products = scrape_category_products(cat_url)
    print(f"   → Found {len(products)} products")

    for p_idx, (pname, purl) in enumerate(products, start=1):
        if stop_requested:
            break

        total_scraped += 1
        print(f"      ({p_idx}/{len(products)}) Scraping: {pname}")

        product_info = scrape_product_details(purl)
        data.append(product_info)

        # Auto-save after every 50 products
        if total_scraped % SAVE_INTERVAL == 0:
            save_data()

print("\n✅ Finished scraping all categories!")
save_data()
driver.quit()
print("🎉 All done!")
