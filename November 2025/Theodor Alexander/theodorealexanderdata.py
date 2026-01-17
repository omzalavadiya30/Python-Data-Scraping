# coding: utf-8
import time
import random
import pandas as pd
from urllib.parse import urljoin
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    StaleElementReferenceException,
    ElementClickInterceptedException
)

# ---------------- CONFIG ----------------
BASE_URL = "https://www.theodorealexander.com"
OUTPUT_FILE = "theodorealexander_products.xlsx"
SAVE_INTERVAL= 50

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_argument("--disable-notifications")
chrome_options.add_argument("--no-sandbox")

driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 20)

# ---------------- HELPER FUNCTION ----------------
def scroll_and_wait(scroll_pause=1.8, step=900):
    """Scroll slowly through the page to trigger lazy loading."""
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script(f"window.scrollBy(0, {step});")
        time.sleep(scroll_pause)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height
    time.sleep(1)

def collect_products(main_cat, sub_cat):
    """Collect all product URLs for one subcategory (returns list of urls)."""
    all_links = []
    seen_links = set()

    page_num = 1
    prev_total = 0

    while True:
        scroll_and_wait()

        try:
            products = driver.find_elements(By.CSS_SELECTOR, "#productListDiv div.product a.productImage")
        except StaleElementReferenceException:
            products = []

        added = 0

        for p in products:
            try:
                href = p.get_attribute("href")
                name = p.get_attribute("title") or ""

                # Handle relative links (Collections pages)
                if href and href.startswith("/"):
                    href = urljoin(BASE_URL, href)

                if href and href not in seen_links:
                    seen_links.add(href)
                    all_links.append({"product_url": href, "product_name_guess": name.strip()})
                    added += 1
            except StaleElementReferenceException:
                continue

        total_now = len(all_links)
        print(f"      📄 Page {page_num}: {added} new products (Total so far: {total_now})")

        # stop if no new products
        if added == 0 or total_now == prev_total:
            print(f"      🛑 No new products detected — pagination ended for {sub_cat}")
            break

        prev_total = total_now

        # --- Safe Pagination Click ---
        try:
            next_btn = driver.find_element(By.CSS_SELECTOR, "button[data-page='next']")
            btn_disabled = ("disabled" in (next_btn.get_attribute("class") or "").lower()) or (not next_btn.is_enabled())
            if btn_disabled:
                print(f"      🛑 Next button disabled — end of pages for {sub_cat}")
                break

            # Scroll into view and wait for clickability
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", next_btn)
            time.sleep(1)

            try:
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-page='next']")))
                next_btn.click()
            except ElementClickInterceptedException:
                # Fallback: JS click when intercepted
                print("⚠️  Next button click intercepted — retrying with JS click")
                driver.execute_script("arguments[0].click();", next_btn)
            except Exception as e:
                print(f"⚠️  Click failed ({type(e).__name__}) — retrying once...")
                driver.execute_script("arguments[0].click();", next_btn)

            page_num += 1
            time.sleep(2.5)

            # Wait for products to reload
            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located(
                    (By.CSS_SELECTOR, "#productListDiv div.product a.productImage")
                )
            )
            time.sleep(random.uniform(1.5, 3.5))  # random human delay

        except (NoSuchElementException, TimeoutException):
            print(f"      🛑 No Next button — finished {sub_cat}")
            break
        except Exception as e:
            print(f"⚠️ Pagination click failed ({type(e).__name__}): {e}")
            break

    return all_links


def safe_find_text(by, selector, default=""):
    try:
        return driver.find_element(by, selector).text.strip()
    except Exception:
        return default


def safe_find_attr(by, selector, attr="href", default=""):
    try:
        return driver.find_element(by, selector).get_attribute(attr) or default
    except Exception:
        return default

def get_detail_label_value(label_keyword):
    """
    In the Details tab (#nav-detail), find the li where the title contains label_keyword
    and return the inner text or html of the value span (col-xl-8).
    """
    try:
        items = driver.find_elements(By.CSS_SELECTOR, "#nav-detail .product_tab_content_detail-ul .product_tab_content_detail-li")
        for li in items:
            try:
                title_el = li.find_element(By.CSS_SELECTOR, ".product_tab_content_detail-title")
                title_text = title_el.text.strip()
                if label_keyword.lower() in title_text.lower():
                    # value span might be .col-xl-8 or .product_tab_content_detail-content
                    try:
                        val = li.find_element(By.CSS_SELECTOR, "span.col-xl-8")
                        return val.get_attribute("innerHTML").strip()
                    except Exception:
                        # fallback: return whole li innerHTML
                        return li.get_attribute("innerHTML").strip()
            except Exception:
                continue
    except Exception:
        pass
    return ""


def get_material_value(label_text):
    """
    Finds the text value of a detail field (like Main Materials, Finish Materials)
    inside #nav-detail using flexible matching.
    """
    try:
        detail_items = driver.find_elements(By.CSS_SELECTOR, "#nav-detail li.product_tab_content_detail-li")
        for item in detail_items:
            try:
                label = item.find_element(By.CSS_SELECTOR, ".product_tab_content_detail-title").text.strip().lower()
                if label_text.lower().rstrip(":") in label:
                    # Direct <span class="col-xl-8">
                    try:
                        value_el = item.find_element(By.CSS_SELECTOR, "span.col-xl-8")
                        value_text = value_el.text.strip()
                        if value_text:
                            return value_text
                    except Exception:
                        pass

                    # Nested inside <div class="product_tab_content_detail-content">
                    try:
                        value_el = item.find_element(By.CSS_SELECTOR, ".product_tab_content_detail-content .col-xl-8")
                        value_text = value_el.text.strip()
                        if value_text:
                            return value_text
                    except Exception:
                        pass

                    # Fallback: full text minus label
                    try:
                        full_text = item.text.strip()
                        value_text = full_text.replace(label, "", 1).strip(": ").strip()
                        if value_text:
                            return value_text
                    except Exception:
                        pass
            except Exception:
                continue
    except Exception:
        pass
    return "N/A"

# ---------------- CATEGORY LINKS ----------------
categories = {
    "Living": {
        "Accent Chairs": "https://www.theodorealexander.com/item/category/type/value/accent-chairs",
        "Accent Tables": "https://www.theodorealexander.com/item/category/type/value/accent-tables",
        "Center Tables": "https://www.theodorealexander.com/item/category/type/value/center-tables",
        "Cocktail Tables": "https://www.theodorealexander.com/item/category/type/value/cocktail-tables",
        "Console Tables": "https://www.theodorealexander.com/item/category/type/value/console-tables",
        "Game Tables": "https://www.theodorealexander.com/item/category/type/value/game-tables",
        "Media Cabinets": "https://www.theodorealexander.com/item/category/type/value/media-cabinets",
        "Ottomans & Stools": "https://www.theodorealexander.com/item/category/type/value/ottomans--stools",
        "Sectionals": "https://www.theodorealexander.com/item/category/type/value/sectionals",
        "Side Tables": "https://www.theodorealexander.com/item/category/type/value/side-tables",
        "Sofas & Settees": "https://www.theodorealexander.com/item/category/type/value/sofas--settees",
        "Storage Cabinets": "https://www.theodorealexander.com/item/category/type/value/storage-cabinets"
    },
    "Dining": {
        "Bar & Counter Stools": "https://www.theodorealexander.com/item/category/type/value/bar--counter-stools",
        "Bar & Pub Tables": "https://www.theodorealexander.com/item/category/type/value/bar--pub-tables",
        "Bar Carts & Cabinets": "https://www.theodorealexander.com/item/category/type/value/bar-carts--cabinets",
        "China & Curio Cabinets": "https://www.theodorealexander.com/item/category/type/value/china--curio-cabinets",
        "Dining Chairs": "https://www.theodorealexander.com/item/category/type/value/dining-chairs",
        "Rectangular & Oval Dining Table": "https://www.theodorealexander.com/item/category/type/value/rectangular--oval-dining-table",
        "Round Dining Tables": "https://www.theodorealexander.com/item/category/type/value/round-dining-tables",
        "Sideboards & Buffets": "https://www.theodorealexander.com/item/category/type/value/sideboards--buffets"
    },
    "Bed": {
        "Beds": "https://www.theodorealexander.com/item/category/type/value/beds",
        "Benches": "https://www.theodorealexander.com/item/category/type/value/benches",
        "Dressers & Chests": "https://www.theodorealexander.com/item/category/type/value/dressers--chests",
        "Nightstands": "https://www.theodorealexander.com/item/category/type/value/nightstands",
        "Storage": "https://www.theodorealexander.com/item/category/type/value/storage",
        "Vanity Tables": "https://www.theodorealexander.com/item/category/type/value/vanity-tables"
    },
    "Office": {
        "Bookcases & Etageres": "https://www.theodorealexander.com/item/category/type/value/bookcases--etageres",
        "Desk Chairs": "https://www.theodorealexander.com/item/category/type/value/desk-chairs",
        "Desks & Bureauxs": "https://www.theodorealexander.com/item/category/type/value/desks--bureauxs"
    },
    "Lighting": {
        "Ceiling Lighting": "https://www.theodorealexander.com/item/category/type/value/ceiling-lighting",
        "Floor Lighting": "https://www.theodorealexander.com/item/category/type/value/floor-lighting",
        "Table Lighting": "https://www.theodorealexander.com/item/category/type/value/table-lighting"
    },
    "Decor": {
        "Free Standing Accessories": "https://www.theodorealexander.com/item/category/type/value/free-standing-accessories",
        "Mirrors": "https://www.theodorealexander.com/item/category/type/value/mirrors",
        "Table Top Accessories": "https://www.theodorealexander.com/item/category/type/value/table-top-accessories",
        "Wall Art": "https://www.theodorealexander.com/item/category/type/value/wall-art"
    },
    "Collections": {
        "Alexa Hampton": "https://www.theodorealexander.com/item/category/collection/value/the-alexa-hampton-collection",
        "Althorp - Victory Oak": "https://www.theodorealexander.com/item/category/collection/value/althorp--victory-oak",
        "Althorp Living History": "https://www.theodorealexander.com/item/category/collection/value/althorp-living-history",
        "Brooksby": "https://www.theodorealexander.com/item/category/collection/value/brooksby",
        "Brushwork New": "https://www.theodorealexander.com/item/category/collection/value/brushwork",
        "Castle Bromwich": "https://www.theodorealexander.com/item/category/collection/value/castle-bromwich",
        "Dorchester": "https://www.theodorealexander.com/item/category/collection/value/dorchester",
        "Morning Room": "https://www.theodorealexander.com/item/category/collection/value/morning-room",
        "Sloane": "https://www.theodorealexander.com/item/category/collection/value/sloane",
        "Spencer London": "https://www.theodorealexander.com/item/category/collection/value/spencer-london",
        "Stephen Church": "https://www.theodorealexander.com/item/category/collection/value/the-stephen-church-collection",
        "Tavel": "https://www.theodorealexander.com/item/category/collection/value/the-tavel-collection",
        "Arakan New": "https://www.theodorealexander.com/item/category/collection/value/arakan",
        "Bouquet New": "https://www.theodorealexander.com/item/category/collection/value/bouquet",
        "Catalina": "https://www.theodorealexander.com/item/category/collection/value/catalina",
        "Horizon": "https://www.theodorealexander.com/item/category/collection/value/horizon",
        "Hudson": "https://www.theodorealexander.com/item/category/collection/value/hudson-collection",
        "Isola": "https://www.theodorealexander.com/item/category/collection/value/the-isola-collection",
        "Judith Leiber Couture": "https://www.theodorealexander.com/item/category/collection/value/judith-leiber-couture",
        "Kesden": "https://www.theodorealexander.com/item/category/collection/value/kesden-collection",
        "Luna": "https://www.theodorealexander.com/item/category/collection/value/luna",
        "Maxwell": "https://www.theodorealexander.com/item/category/collection/value/maxwell",
        "Origins": "https://www.theodorealexander.com/item/category/collection/value/origins",
        "Panos New": "https://www.theodorealexander.com/item/category/collection/value/panos",
        "Repose": "https://www.theodorealexander.com/item/category/collection/value/repose",
        "Rome": "https://www.theodorealexander.com/item/category/collection/value/rome",
        "Spencer ST. James": "https://www.theodorealexander.com/item/category/collection/value/spencer-st-james",
        "Urbane": "https://www.theodorealexander.com/item/category/collection/value/urbane",
        "Balboa": "https://www.theodorealexander.com/item/category/collection/value/balboa",
        "Breeze": "https://www.theodorealexander.com/item/category/collection/value/breeze-collection",
        "Echoes": "https://www.theodorealexander.com/item/category/collection/value/the-echoes-collection",
        "Essence": "https://www.theodorealexander.com/item/category/collection/value/essence-collection",
        "LIDO": "https://www.theodorealexander.com/item/category/collection/value/lido-collection",
        "Montauk": "https://www.theodorealexander.com/item/category/collection/value/montauk",
        "NOVA": "https://www.theodorealexander.com/item/category/collection/value/nova-collection",
        "Surrey": "https://www.theodorealexander.com/item/category/collection/value/surrey",
        "Accessories": "https://www.theodorealexander.com/item/category/collection/value/accessories",
        "Art by TA": "https://www.theodorealexander.com/item/category/collection/value/art-by-ta",
        "Floored": "https://www.theodorealexander.com/item/category/collection/value/floored",
        "Marlborough by Alexa Hampton New": "https://www.theodorealexander.com/item/category/collection/value/marlborough-by-alexa-hampton",
        "Seated": "https://www.theodorealexander.com/item/category/collection/value/seated",
        "Spencer Coronet New": "https://www.theodorealexander.com/item/category/collection/value/spencer-coronet",
        "TA Artistry": "https://www.theodorealexander.com/item/category/collection/value/ta-artistry",
        "TA Illuminations": "https://www.theodorealexander.com/item/category/collection/value/ta-illuminations",
        "TA Originals": "https://www.theodorealexander.com/item/category/collection/value/ta-originals",
        "TA Studio": "https://www.theodorealexander.com/item/category/collection/value/ta-studio",
        "TAilor Fit": "https://www.theodorealexander.com/tailorfit",
        "THEO by Theodore Alexander": "https://www.theodorealexander.com/item/category/collection/value/theo-by-theodore-alexander",
        "Upholstery": "https://www.theodorealexander.com/item/category/collection/value/upholstery"
    }
}

# ---------------- SCRAPING ----------------
collected_links = []

print("Collecting product URLs (robust pagination with scroll + stop detection)...\n")

for main_cat, subs in categories.items():
    print(f"\n🔹 Main Category: {main_cat}")
    for sub_cat, url in subs.items():
        print(f"   🟢 Subcategory: {sub_cat}")
        driver.get(url)
        time.sleep(4)

        try:
            wait.until(EC.presence_of_all_elements_located(
                (By.CSS_SELECTOR, "#productListDiv div.product a.productImage")
            ))
        except TimeoutException:
            print(f"      ⚠️ No products found for {sub_cat}")
            continue

        sub_links = collect_products(main_cat, sub_cat)
        collected_links.extend(sub_links)

        print(f"      ✅ {len(sub_links)} total products collected from {sub_cat}")

# Normalize unique product URLs (preserve order)
unique_urls = []
seen_urls = set()
for entry in collected_links:
    u = entry.get("product_url")
    if u and u not in seen_urls:
        seen_urls.add(u)
        unique_urls.append(u)

print(f"\n🔎 Total unique product pages to scrape: {len(unique_urls)}")

# ---------------- PER-PRODUCT SCRAP ----------------
final_data = []

def save_progress():
    if final_data:
        df = pd.DataFrame(final_data)
        df.to_excel(OUTPUT_FILE, index=False)
        print(f"\n💾 Progress saved! Total products saved: {len(final_data)}")
try:
    for idx, product_url in enumerate(unique_urls, start=1):
        print(f"\n[{idx}/{len(unique_urls)}] Scraping: {product_url}")
        try:
            driver.get(product_url)
        except Exception as e:
            print("  ⚠️ Failed to load page:", e)
            continue

        # Wait for main name or sku to appear (or timeout)
        try:
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".product_detail_info_name")))
        except TimeoutException:
            print("  ⚠️ product name didn't appear quickly; continuing with what we can.")

        time.sleep(1.0)

        # Basic fields
        category_field = ""
        # Try to extract Room / Type anchors inside the details section
        try:
            # click Details tab to ensure it's active (usually active by default)
            try:
                detail_tab_btn = driver.find_element(By.CSS_SELECTOR, "#nav-detail-tab")
                driver.execute_script("arguments[0].click();", detail_tab_btn)
                time.sleep(0.4)
            except Exception:
                pass

            # Find 'Room / Type' value and format "Living Room >> Accent Chairs"
            try:
                items = driver.find_elements(By.CSS_SELECTOR, "#nav-detail .product_tab_content_detail-li")
                room_val = ""
                for li in items:
                    try:
                        title_el = li.find_element(By.CSS_SELECTOR, ".product_tab_content_detail-title")
                        if "Room / Type" in title_el.text:
                            val_span = li.find_element(By.CSS_SELECTOR, "span.col-xl-8")
                            # collect anchor texts
                            anchors = val_span.find_elements(By.TAG_NAME, "a")
                            anchor_texts = [a.text.strip() for a in anchors if a.text.strip()]
                            if anchor_texts:
                                category_field = " >> ".join(anchor_texts)
                                room_val= val_span.text.strip()
                            else:
                                category_field = val_span.text.strip()
                            break
                    except Exception:
                        continue
            except Exception:
                category_field = ""
        except Exception:
            category_field = ""

        # Product URL
        url_field = product_url

        # Product Name
        product_name = safe_find_text(By.CSS_SELECTOR, ".product_detail_info_name", "")

        # SKU
        sku = safe_find_text(By.CSS_SELECTOR, ".inner-sku", "")

        # Brand (fixed)
        brand = "Theodore Alexander"

        # DETAILS_FULL_DESCRIPTION (full html of #nav-detail)
        details_full_html = ""
        try:
            # ensure Details tab is active
            try:
                driver.execute_script("document.querySelector('#nav-detail-tab')?.click();")
                time.sleep(0.4)
            except Exception:
                pass

            details_el = driver.find_element(By.CSS_SELECTOR, "#nav-detail")
            details_full_html = details_el.get_attribute("innerHTML").strip()
        except Exception:
            details_full_html = ""

        # Collection
        collection_html = get_detail_label_value("Collection:")
        # If it's an anchor, get text only
        collection_text = ""
        if collection_html:
            # strip tags if any - simple approach: if there is an <a> tag, get its text via selenium
            try:
                coll_el = driver.find_element(By.XPATH, "//div[@id='nav-detail']//span[contains(.,'Collection')]/following::span[1]//a")
                collection_text = coll_el.text.strip()
            except Exception:
                # fallback: plain text from html
                collection_text = (collection_html.replace("\n", " ").strip())
        else:
            collection_text = ""

        # Use new logic for materials
        main_material = get_material_value("Main Materials")
        finish_material = get_material_value("Finish Materials")

        # Net Weight / Gross Weight
        net_weight = ""
        gross_weight = ""
        try:
            # find the li where title contains 'Net Weight' and 'Gross Weight'
            lis = driver.find_elements(By.CSS_SELECTOR, "#nav-detail .product_tab_content_detail-li")
            for li in lis:
                try:
                    title_el = li.find_element(By.CSS_SELECTOR, ".product_tab_content_detail-title")
                    t = title_el.text.strip()
                    if "Net Weight" in t:
                        # get the value
                        try:
                            val = li.find_element(By.CSS_SELECTOR, ".product_tab_content_detail-content")
                            net_weight = val.text.strip()
                        except Exception:
                            net_weight = li.text.replace(t, "").strip()
                    if "Gross Weight" in t:
                        try:
                            val = li.find_element(By.CSS_SELECTOR, ".product_tab_content_detail-content")
                            gross_weight = val.text.strip()
                        except Exception:
                            gross_weight = li.text.replace(t, "").strip()
                except Exception:
                    continue
        except Exception:
            pass

        # Dimensions: Attr_width_in, Attr_depth_in, Attr_height_in, Attr_width_cm, Attr_depth_cm, Attr_height_cm
        Attr_width_In = Attr_depth_In = Attr_height_In = ""
        Attr_width_Cm = Attr_depth_Cm = Attr_height_Cm = ""
        try:
            table_rows = driver.find_elements(By.CSS_SELECTOR, ".product_detail_info_dimension table.tableDimension tbody tr")
            for tr in table_rows:
                try:
                    unit = tr.find_element(By.CSS_SELECTOR, "th").text.strip().lower()
                    tds = tr.find_elements(By.TAG_NAME, "td")
                    if len(tds) >= 3:
                        w = tds[0].text.strip()
                        d = tds[1].text.strip()
                        h = tds[2].text.strip()
                        if unit.startswith("in"):
                            Attr_width_In = w
                            Attr_depth_In = d
                            Attr_height_In = h
                        elif unit.startswith("cm"):
                            Attr_width_Cm = w
                            Attr_depth_Cm = d
                            Attr_height_Cm = h
                except Exception:
                    continue
        except Exception:
            pass

        # Description (short) - innerHTML of .product_detail_info_description
        description_html = ""
        try:
            desc_el = driver.find_element(By.CSS_SELECTOR, ".product_detail_info_description")
            description_html = desc_el.text.strip()
        except Exception:
            description_html = ""

        # Full Description - full html code of .product_detail_info
        full_description_html = ""
        try:
            pd_info = driver.find_element(By.CSS_SELECTOR, ".product_detail_info")
            full_description_html = pd_info.get_attribute("innerHTML").strip()
        except Exception:
            full_description_html = ""

        # Furniture Care full html: requires clicking the nav-care tab
        furniture_care_html = ""
        try:
            # Click the care tab
            try:
                care_btn = driver.find_element(By.CSS_SELECTOR, "#nav-care-tab")
                driver.execute_script("arguments[0].click();", care_btn)
                time.sleep(0.6)
            except Exception:
                # maybe not present - try directly
                pass

            try:
                care_el = driver.find_element(By.CSS_SELECTOR, "#nav-care")
                furniture_care_html = care_el.get_attribute("innerHTML").strip()
            except Exception:
                furniture_care_html = ""
        except Exception:
            furniture_care_html = ""

        # Resources: click resources tab and extract specific pdf links + full html
        resources_full_html = ""
        com_form_url = ""
        tearsheet_url = ""
        frame_availability_url = ""
        alexa_hampton_catalog_url = ""
        fabric_trim_availability_url = ""
        try:
            # Click resources tab (note id nav-resouces)
            try:
                res_btn = driver.find_element(By.CSS_SELECTOR, "#nav-resouces-tab")
                driver.execute_script("arguments[0].click();", res_btn)
                time.sleep(0.6)
            except Exception:
                pass

            try:
                res_el = driver.find_element(By.CSS_SELECTOR, "#nav-resouces")
                resources_full_html = res_el.get_attribute("innerHTML").strip()

                # find links under resources
                download_links = res_el.find_elements(By.CSS_SELECTOR, "a.product_tab_content_download-a")
                for a in download_links:
                    try:
                        key_el = a.find_element(By.CSS_SELECTOR, ".product_tab_content_download_key-uppercase")
                        key_text = key_el.text.strip().lower()
                        href = a.get_attribute("href") or ""
                        if "com form" in key_text or "com-form" in key_text or "com" == key_text.lower().strip():
                            com_form_url = href
                        elif "tear sheet" in key_text or "tear sheet" in key_text:
                            tearsheet_url = href
                        elif "frame availability" in key_text:
                            frame_availability_url = href
                        elif "the alexa hampton catalog" in key_text or "alexa hampton" in key_text.lower():
                            alexa_hampton_catalog_url = href
                        elif "fabric and trim availability" in key_text or "fabric and trim" in key_text.lower():
                            fabric_trim_availability_url = href
                    except Exception:
                        # fallback try to inspect the inner text
                        try:
                            href = a.get_attribute("href") or ""
                            txt = a.text.lower()
                            if "com form" in txt and not com_form_url:
                                com_form_url = href
                            if "tear" in txt and not tearsheet_url:
                                tearsheet_url = href
                            if "frame" in txt and not frame_availability_url:
                                frame_availability_url = href
                            if "alexa hampton" in txt and not alexa_hampton_catalog_url:
                                alexa_hampton_catalog_url = href
                            if "fabric and trim" in txt and not fabric_trim_availability_url:
                                fabric_trim_availability_url = href
                        except Exception:
                            continue
            except Exception:
                pass
        except Exception:
            pass

        # Images: extract up to 4 images from '.smv-scroll img' (src)
        image_urls = []
        try:
            img_els = driver.find_elements(By.CSS_SELECTOR, ".smv-scroll img")
            for img in img_els:
                try:
                    s = img.get_attribute("src") or img.get_attribute("data-src") or ""
                    if s and s not in image_urls:
                        image_urls.append(s)
                    if len(image_urls) >= 4:
                        break
                except Exception:
                    continue
        except Exception:
            image_urls = []

        # Assemble row in the required order:
        row = {
            "Category": category_field,
            "Product URL": url_field,
            "Product Name": product_name,
            "SKU": sku,
            "Brand": brand,
            "Collection": collection_text,
            "Room/type": room_val,  # same as Category field per requirement
            "Main Material": main_material,
            "Finish Material": finish_material,
            "Net Weight": net_weight,
            "Gross Weight": gross_weight,
            "Attr_Width_In": Attr_width_In,
            "Attr_Depth_In": Attr_depth_In,
            "Attr_Height_In": Attr_height_In,
            "Attr_Width_Cm": Attr_width_Cm,
            "Attr_Depth_Cm": Attr_depth_Cm,
            "Attr_Height_Cm": Attr_height_Cm,
            "COM Form": com_form_url,
            "Tearsheet": tearsheet_url,
            "Frame Availability": frame_availability_url,
            "The Alexa Hampton Catalog": alexa_hampton_catalog_url,
            "Fabric and Trim Availability": fabric_trim_availability_url,
            "Description": description_html,
            "DETAILS_FULL_DESCRIPTION": details_full_html,
            "Full Description": full_description_html,
            "Furniture Care full html": furniture_care_html,
            "Resources_full_description": resources_full_html,
            "Image1": image_urls[0] if len(image_urls) > 0 else "",
            "Image2": image_urls[1] if len(image_urls) > 1 else "",
            "Image3": image_urls[2] if len(image_urls) > 2 else "",
            "Image4": image_urls[3] if len(image_urls) > 3 else "",
        }

        # ✅ Replace any missing or empty values with "N/A"
        for key, value in row.items():
            if not value or str(value).strip() == "":
                row[key] = "N/A"


        final_data.append(row)
        if idx % SAVE_INTERVAL == 0:
            save_progress()

        # small polite pause between pages
        time.sleep(0.6)
except KeyboardInterrupt:
    print("\n⚠️ Script interrupted by user.")
    save_progress()

except Exception as e:
    print(f"\n❌ Unexpected error: {e}")
    save_progress()

finally:
    # Save final progress
    save_progress()
    driver.quit()
    print(f"\n✅ Done! Total products collected: {len(final_data)}")