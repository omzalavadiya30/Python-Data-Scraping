import os
import time
import json
import signal
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ---------------- CONFIG ----------------
OUTPUT_FILE = "point1920_product_details.xlsx"
BACKUP_FILE = "point1920_backup.json"
SAVE_INTERVAL = 50  # Auto-save after every 50 products

CATEGORIES = {
    "Products": [
        "https://www.point1920.com/shop-online-6212",
        "https://www.point1920.com/chairs",
        "https://www.point1920.com/armchairs",
        "https://www.point1920.com/stools",
        "https://www.point1920.com/tables",
        "https://www.point1920.com/designer-sofas",
        "https://www.point1920.com/puffs",
        "https://www.point1920.com/cushions-quadrants",
        "https://www.point1920.com/sunbeds",
        "https://www.point1920.com/umbrellas",
        "https://www.point1920.com/swings",
        "https://www.point1920.com/planters",
        "https://www.point1920.com/lamps",
    ],
    "Collections": [
        "https://www.point1920.com/curio-collection",
        "https://www.point1920.com/amba-collection",
        "https://www.point1920.com/zuma-collection",
        "https://www.point1920.com/legacy_collection",
        "https://www.point1920.com/kubik_collection",
        "https://www.point1920.com/neck_collection",
        "https://www.point1920.com/kahn_collection",
        "https://www.point1920.com/origin_collection",
        "https://www.point1920.com/longisland-collection",
        "https://www.point1920.com/city-collection",
        "https://www.point1920.com/lis-collection",
        "https://www.point1920.com/summer-collection",
        "https://www.point1920.com/sunset-collection",
        "https://www.point1920.com/khai-6476",
        "https://www.point1920.com/weave_collection",
        "https://www.point1920.com/arc-collection",
        "https://www.point1920.com/bay-8526",
        "https://www.point1920.com/hamp-collection",
        "https://www.point1920.com/pal-collection",
        "https://www.point1920.com/fup-collection",
        "https://www.point1920.com/round-collection",
        "https://www.point1920.com/colorscompact-collection",
        "https://www.point1920.com/heritage_collection",
        "https://www.point1920.com/paralel-collection",
        "https://www.point1920.com/alga-collection_",
        "https://www.point1920.com/arena-collection",
        "https://www.point1920.com/breda-collection",
        "https://www.point1920.com/brumas-collection",
        "https://www.point1920.com/caddie-collection",
        "https://www.point1920.com/charleston-collection",
        "https://www.point1920.com/havana-collection",
        "https://www.point1920.com/romantic-collection",
        "https://www.point1920.com/sagra-2461"
    ],
    "Projects": [
        "https://www.point1920.com/project-hotel-aethos-mallorca-spain",
        "https://www.point1920.com/project-hotel-barriere-le-normandy-france",
        "https://www.point1920.com/project-casa-tosca-spain",
        "https://www.point1920.com/project-eurostars-casa-anfa-morocco",
        "https://www.point1920.com/project-four-seasons-resort-mallorca",
        "https://www.point1920.com/project-sha-wellness-hotel-mexico",
        "https://www.point1920.com/project-umusic-rooftop-casa-chicote-en",
        "https://www.point1920.com/project-vilalara-grand-hotel-algarve",
        "https://www.point1920.com/categories/779",
        "https://www.point1920.com/project-quinto-sol-house-en",
        "https://www.point1920.com/project-arena-beach-club-en",
        "https://www.point1920.com/project-four-seasons-resort-marrakech",
        "https://www.point1920.com/project-hipotels-mediterraneo-en",
        "https://www.point1920.com/project-hipotels-playa-palma-palace-en",
        "https://www.point1920.com/project-ivey-on-boren-usa",
        "https://www.point1920.com/project-the-view-agadir-en",
        "https://www.point1920.com/project-vocco-stockholm-kista-en",
        "https://www.point1920.com/categories/783",
        "https://www.point1920.com/project-wynwood-haus-en",
        "https://www.point1920.com/la-peer-hotel-hollywood",
        "https://www.point1920.com/softiel-agadir-thalassa-sea-spa",
        "https://www.point1920.com/cosme-hotel",
        "https://www.point1920.com/alys_beach_club_us",
        "https://www.point1920.com/1_hotel_toronto_canada",
        "https://www.point1920.com/anaw-restaurant-morocco",
        "https://www.point1920.com/like-a-palm-tree-restaurant",
        "https://www.point1920.com/hotel-santa-m-briones-spain",
        "https://www.point1920.com/kilroy_sabre_springs_california",
        "https://www.point1920.com/radisson_collection_sevilla_spain",
        "https://www.point1920.com/one_and_only_royal_mirage_resort",
        "https://www.point1920.com/residences-by-armani-casa-us",
        "https://www.point1920.com/blue-mountain-guest-house-myeongdond-south-korea",
        "https://www.point1920.com/blue-palace-marriott-greece",
        "https://www.point1920.com/hotel-el-fuerte-spain",
        "https://www.point1920.com/one_and_only",
        "https://www.point1920.com/sepp-alpine-boutique-hotel-2380",
        "https://www.point1920.com/villa-castanyetes-mallorca-spain",
        "https://www.point1920.com/hilton-taghazout-blue-beach-morocco",
        "https://www.point1920.com/hotel-portocolom-spain",
        "https://www.point1920.com/categories/681",
        "https://www.point1920.com/masia-bellver-restaurant",
        "https://www.point1920.com/grand-sirenis-resort-dominican-republic",
        "https://www.point1920.com/fisher_island_club_miami_us",
        "https://www.point1920.com/kalida-san-pau-spain",
        "https://www.point1920.com/the-langham-hotels-resorts-australia_",
        "https://www.point1920.com/four-points-by-sheraton-algeria",
        "https://www.point1920.com/zela-london-restaurant-uk",
        "https://www.point1920.com/pleta-de-mar-hotel-spain-",
        "https://www.point1920.com/punta_de_mar_spain",
        "https://www.point1920.com/les-deux-tours-hotel-morocco",
        "https://www.point1920.com/chi-the-spa-shangri-la-hotels-paris-france",
        "https://www.point1920.com/aguas-de-ibiza-spain",
        "https://www.point1920.com/hotel-g-san-francisco-us",
        "https://www.point1920.com/hotel-skye-niseko-japan",
        "https://www.point1920.com/galaxy-macau-china__",
        "https://www.point1920.com/ritmo_formentera_spain",
        "https://www.point1920.com/king-street-hotel-",
        "https://www.point1920.com/al-sarab-desert-resort-by-anantara-uae",
        "https://www.point1920.com/alsacea_hotella_u.s",
        "https://www.point1920.com/ricard-camarena-restarurant-spain_",
        "https://www.point1920.com/w-fort-lauderdale-hotel-us",
        "https://www.point1920.com/grand-luxxe-puerto-vallarta-mexico",
        "https://www.point1920.com/bartaccia-hotel__",
        "https://www.point1920.com/grand-plaza-movenpick-media-city-hotel-%C2%A0uae-",
        "https://www.point1920.com/residence-santa-barbara-goleta-by-marriott",
        "https://www.point1920.com/la-ville-hotel_eau",
        "https://www.point1920.com/al-baleed-resort-by-anantara_",
        "https://www.point1920.com/marina-beach-club-espana",
        "https://www.point1920.com/costa-del-sol-hotel-spain",
        "https://www.point1920.com/schwaiger-xinos-spain",
        "https://www.point1920.com/cabo-negro-royal-golf-club-",
        "https://www.point1920.com/four-seasons-the-nam-hai-vietnam",
        "https://www.point1920.com/rixos_hotel_the_palm",
        "https://www.point1920.com/rancho-punta-mita-mexico_",
        "https://www.point1920.com/norida-beach-greece",
        "https://www.point1920.com/amare-marbella-beach-",
        "https://www.point1920.com/hotel-romazzino-a-luxury-collection-hotel-by-marriott-italy",
        "https://www.point1920.com/hyatt-regency-seragaki-island-okinawa_",
        "https://www.point1920.com/lhotel-elan-china-",
        "https://www.point1920.com/marbella-club-spain",
        "https://www.point1920.com/black-tail-us"
    ],
}

# ---------------- SELENIUM SETUP ----------------
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--disable-notifications")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
wait = WebDriverWait(driver, 15)

# ---------------- UTILS ----------------
def safe_save(data):
    """Save data to Excel and JSON backup."""
    try:
        df = pd.DataFrame(data)
        df.to_excel(OUTPUT_FILE, index=False)
        with open(BACKUP_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"💾 Auto-saved {len(data)} products to {OUTPUT_FILE}")
    except Exception as e:
        print(f"⚠️ Save failed: {e}")

def graceful_exit(signum=None, frame=None):
    print("\n🛑 Graceful exit triggered. Saving progress before quitting...")
    safe_save(all_data)
    driver.quit()
    exit(0)

signal.signal(signal.SIGINT, graceful_exit)
signal.signal(signal.SIGTERM, graceful_exit)

# ---------------- SCRAPE FUNCTIONS ----------------
def get_product_links(category_url):
    """Extract all product links from a category page (supports lazy loading)."""
    try:
        driver.get(category_url)
        time.sleep(3)
        SCROLL_PAUSE_TIME = 2
        product_urls = set()
        last_height = driver.execute_script("return document.body.scrollHeight")
        stable_rounds = 0

        while True:
            # Grab all product URLs visible so far
            elements = driver.find_elements(By.CSS_SELECTOR, "a.product-item-link, div.col.col-product-list a")
            for elem in elements:
                href = elem.get_attribute("href")
                if href:
                    product_urls.add(href)

            # Scroll down
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(SCROLL_PAUSE_TIME)
            new_height = driver.execute_script("return document.body.scrollHeight")

            # If no change in height after several tries → stop
            if new_height == last_height:
                stable_rounds += 1
            else:
                stable_rounds = 0
                last_height = new_height

            if stable_rounds >= 3:
                break

        print(f"✅ Found {len(product_urls)} products in: {category_url}")
        return list(product_urls)

    except Exception as e:
        print(f"⚠️ Failed to load product links from {category_url}: {e}")
        return []

# def extract_product_details(product_url):
#     """Extract product details from detail page."""
#     driver.get(product_url)
#     wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "body")))
#     time.sleep(2)

#     def safe(selector, attr=None, default=""):
#         try:
#             el = driver.find_element(By.CSS_SELECTOR, selector)
#             if attr:
#                 val = el.get_attribute(attr)
#                 return val.strip() if isinstance(val, str) else default
#             return el.text.strip()
#         except:
#             return default

#     # ----- CATEGORY FIX -----
#     product_category = ""
#     try:
#         # Wait until categoryPath anchor is present
#         WebDriverWait(driver, 10).until(
#             EC.presence_of_element_located((By.CSS_SELECTOR, "a.path.categoryPath"))
#         )
#         cat_el = driver.find_element(By.CSS_SELECTOR, "a.path.categoryPath")
#         html_text = cat_el.get_attribute("innerHTML")

#         # Extract only the text from the span with itemprop='name'
#         try:
#             category_el = cat_el.find_element(By.CSS_SELECTOR, "span[itemprop='name']")
#             product_category = category_el.text.strip()
#         except:
#             # fallback: try parsing manually
#             if "itemprop=\"name\"" in html_text:
#                 product_category = html_text.split("itemprop=\"name\">")[-1].split("</span>")[0].strip()
#             else:
#                 product_category = cat_el.text.replace("Back to", "").strip()

#         # clean unwanted prefix if still exists
#         product_category = product_category.replace("Back to", "").strip()
#     except:
#         product_category = ""

#     # ----- BRAND / COLLECTION TEXT & PRODUCT NAME -----
#     # product-brand anchor could be without class 'active-anchor' — use generic selector
#     brand_text = ""
#     try:
#         brand_el = driver.find_element(By.CSS_SELECTOR, "div.product-brand a")
#         brand_text = brand_el.text.replace("\xa0", " ").strip()
#     except:
#         brand_text = ""

#     product_name_raw = safe("h1.product-h1")
#     # Normalize and Title-case the pieces (Curio Collection -> Bar Stool)
#     def normalize_label(s):
#         return " ".join(s.split()).title() if s else ""

#     brand_norm = normalize_label(brand_text)
#     product_norm = normalize_label(product_name_raw)
#     name = f"{brand_norm} -> {product_norm}" if brand_norm else product_norm

#     # SKU
#     sku = safe("div.product-sku").replace("SKU:", "").strip()

#     # Brand fixed
#     brand = "Point Outdoor Living"

#     # SHORT DESCRIPTION (text before "More information")
#     short_desc = ""
#     try:
#         sd = driver.find_element(By.CSS_SELECTOR, "div.product-short-description")
#         text = sd.get_attribute("innerText") or sd.text
#         short_desc = text.split("More information")[0].strip()
#     except:
#         short_desc = ""

#     # FULL DESCRIPTION (outer HTML)
#     full_desc_html = ""
#     try:
#         ld = driver.find_element(By.CSS_SELECTOR, "div.product-long-description")
#         full_desc_html = ld.get_attribute("outerHTML") or ""
#     except:
#         full_desc_html = ""

#     # IMAGES (main + additional)
#     images = []
#     try:
#         main_img = driver.find_element(By.CSS_SELECTOR, "div.gallery-item-count img").get_attribute("src")
#         if main_img:
#             images.append(main_img)
#     except:
#         pass

#     try:
#         extra_imgs = driver.find_elements(By.CSS_SELECTOR, "section.additional-images-product-slider img")
#         for img in extra_imgs:
#             src = img.get_attribute("src")
#             if src:
#                 images.append(src)
#     except:
#         pass

#     image_dict = {f"Image{i+1}": images[i] if i < len(images) else "" for i in range(4)}

#     return {
#         "Category": product_category,
#         "Product URL": product_url,
#         "Name": name,
#         "SKU": sku,
#         "Brand": brand,
#         "Short Description": short_desc,
#         "Full Description HTML": full_desc_html,
#         **image_dict,
#     }

def extract_product_details(product_url):
    """Extract product details from detail page."""
    driver.get(product_url)
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "body")))
    time.sleep(2)

    def safe(selector, attr=None, default=""):
        try:
            el = driver.find_element(By.CSS_SELECTOR, selector)
            if attr:
                val = el.get_attribute(attr)
                return val.strip() if isinstance(val, str) else default
            return el.text.strip()
        except:
            return default

    # ----- CATEGORY FIX -----
    product_category = ""
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "a.path.categoryPath"))
        )
        cat_el = driver.find_element(By.CSS_SELECTOR, "a.path.categoryPath")
        html_text = cat_el.get_attribute("innerHTML")

        try:
            category_el = cat_el.find_element(By.CSS_SELECTOR, "span[itemprop='name']")
            product_category = category_el.text.strip()
        except:
            if "itemprop=\"name\"" in html_text:
                product_category = html_text.split("itemprop=\"name\">")[-1].split("</span>")[0].strip()
            else:
                product_category = cat_el.text.replace("Back to", "").strip()
        product_category = product_category.replace("Back to", "").strip()
    except:
        product_category = ""

    # ----- BRAND / COLLECTION TEXT & PRODUCT NAME -----
    brand_text = ""
    try:
        brand_el = driver.find_element(By.CSS_SELECTOR, "div.product-brand a")
        brand_text = brand_el.text.replace("\xa0", " ").strip()
    except:
        brand_text = ""

    product_name_raw = safe("h1.product-h1")

    def normalize_label(s):
        return " ".join(s.split()).title() if s else ""

    brand_norm = normalize_label(brand_text)
    product_norm = normalize_label(product_name_raw)
    name = f"{brand_norm} -> {product_norm}" if brand_norm else product_norm

    sku = safe("div.product-sku").replace("SKU:", "").strip()
    brand = "Point Outdoor Living"

    # SHORT DESCRIPTION
    short_desc = ""
    try:
        sd = driver.find_element(By.CSS_SELECTOR, "div.product-short-description")
        text = sd.get_attribute("innerText") or sd.text
        short_desc = text.split("More information")[0].strip()
    except:
        short_desc = ""

    # FULL DESCRIPTION (HTML)
    full_desc_html = ""
    try:
        ld = driver.find_element(By.CSS_SELECTOR, "div.product-long-description")
        full_desc_html = ld.get_attribute("outerHTML") or ""
    except:
        full_desc_html = ""

    # IMAGE EXTRACTION
    images = []
    try:
        main_img = driver.find_element(By.CSS_SELECTOR, "div.gallery-item-count img").get_attribute("src")
        if main_img:
            images.append(main_img)
    except:
        pass

    try:
        extra_imgs = driver.find_elements(By.CSS_SELECTOR, "section.additional-images-product-slider img")
        for img in extra_imgs:
            src = img.get_attribute("src")
            if src:
                images.append(src)
    except:
        pass

    image_dict = {f"Image{i+1}": images[i] if i < len(images) else "" for i in range(4)}

    # ================== NEW ATTRIBUTE EXTRACTION ==================
    attr_color = attr_length = attr_width = attr_height = ""
    attr_weight = attr_seatheight = attr_volume = attr_fabrics = ""

    try:
        paragraphs = driver.find_elements(By.CSS_SELECTOR, "div.product-long-description p")
        for p in paragraphs:
            text = p.text.strip()

            # Color or Fabrics
            if text.lower().startswith("colour:"):
                val = text.split(":", 1)[1].strip()
                # Heuristic: decide if this is fabrics or color
                if any(c.isdigit() for c in val):  # e.g. Colour: 0021, 0022
                    attr_fabrics = val
                else:
                    attr_color = val

            # Length × Width × Height
            elif "length" in text.lower() and "width" in text.lower() and "height" in text.lower():
                try:
                    dims = text.split(":")[1].strip().split("x")
                    if len(dims) >= 3:
                        attr_length = dims[0].strip().split()[0]
                        attr_width = dims[1].strip().split()[0]
                        attr_height = dims[2].strip().split()[0]
                except:
                    pass

            # Weight
            elif text.lower().startswith("weight:"):
                attr_weight = text.split(":", 1)[1].strip()

            # Seat Height
            elif "seat height" in text.lower():
                attr_seatheight = text.split(":", 1)[1].strip()

            # Volume
            elif text.lower().startswith("volume:"):
                attr_volume = text.split(":", 1)[1].strip()
    except Exception as e:
        print(f"⚠️ Attribute extraction failed for {product_url}: {e}")

    # ================== RETURN DICTIONARY ==================
    return {
        "Category": product_category,
        "Product URL": product_url,
        "Name": name,
        "SKU": sku,
        "Brand": brand,
        "Short Description": short_desc,
        "Full Description HTML": full_desc_html,
        "Attr_Color": attr_color,
        "Attr_Length": attr_length,
        "Attr_Width": attr_width,
        "Attr_Height": attr_height,
        "Attr_Weight": attr_weight,
        "Attr_SeatHeight": attr_seatheight,
        "Attr_Volume": attr_volume,
        "Attr_Fabrics": attr_fabrics,
        **image_dict,
    }

# ---------------- MAIN SCRAPER ----------------
all_data = []

# Resume if backup exists
if os.path.exists(BACKUP_FILE):
    try:
        with open(BACKUP_FILE, "r", encoding="utf-8") as f:
            all_data = json.load(f)
        print(f"♻️ Resumed from backup with {len(all_data)} products.")
    except Exception as e:
        print(f"⚠️ Could not load backup: {e}")

for main_cat, urls in CATEGORIES.items():
    print(f"\n=== 🔷 MAIN CATEGORY: {main_cat} ===")

    for category_url in urls:
        print(f"\n🔍 Scraping Category: {category_url}")

        # Step 1: Collect all product URLs for this category
        product_links = get_product_links(category_url)
        total = len(product_links)
        print(f"   → Found {total} product URLs in {category_url}")

        if not product_links:
            print(f"⚠️ No products found in {category_url}")
            continue

        # Step 2: Scrape product details for this category
        for idx, link in enumerate(product_links, start=1):
            print(f"      📦 [{idx}/{len(product_links)}] Scraping → {link}")
            try:
                details = extract_product_details(link)
                all_data.append(details)
            except Exception as e:
                print(f"⚠️ Failed scraping {link}: {e}")

            # Auto-save after every SAVE_INTERVAL products
            if len(all_data) % SAVE_INTERVAL == 0:
                safe_save(all_data)

        # Save after finishing one category
        safe_save(all_data)
        print(f"✅ Finished category: {category_url}")

# Final save
safe_save(all_data)
print("\n✅ Scraping completed successfully.")
driver.quit()