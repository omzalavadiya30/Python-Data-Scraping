import time
import signal
import sys
from urllib.parse import urljoin
from pathlib import Path

import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# ------------- CONFIG -------------
BASE_URL = "https://www.galtechcorp.com/"
OUTPUT_FILE = "galtechcorp_data.xlsx"
CATEGORY_START = "Aluminum"
CATEGORY_END = "Bases"
PAGE_SLEEP = 2.0         # seconds to wait for page load (increase if slow)
MAX_IMAGES = 4
HEADLESS = False         # set True to run headless (if site serves differently in headless, set False)
# ----------------------------------

COLUMNS = [
    "Category",
    "Product URL",
    "Product Name",
    "SKU",
    "Brand",
    "Full Description (HTML)",
    "Attr_SizeText",
    "Attr_SizeImage",
    "Part List (HTML)",
    "Image 1",
    "Image 2",
    "Image 3",
    "Image 4",
]

rows = []           # list of row lists
runtime_driver = None

# ----------------- Utilities -----------------
def abs_url(href):
    if not href:
        return ""
    return href if href.startswith("http") else urljoin(BASE_URL, href)

def new_driver():
    opts = Options()
    if HEADLESS:
        opts.add_argument("--headless=new")
        opts.add_argument("--disable-gpu")
    opts.add_argument("--start-maximized")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    # comment out image-blocking for visual fidelity; uncomment to speed up if needed:
    # opts.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
    return driver

def save_now(reason="manual save"):
    """Write current rows to Excel immediately."""
    try:
        df = pd.DataFrame(rows, columns=COLUMNS)
        df.to_excel(OUTPUT_FILE, index=False)
        print(f"\n💾 Saved {len(df)} rows to '{OUTPUT_FILE}' (reason: {reason})")
    except Exception as e:
        print("❌ Error saving file:", e)

def sigint_handler(signum, frame):
    print("\n\nSIGINT received — saving collected data and exiting...")
    save_now(reason="SIGINT")
    try:
        if runtime_driver:
            runtime_driver.quit()
    except Exception:
        pass
    sys.exit(0)

signal.signal(signal.SIGINT, sigint_handler)

# ----------------- Detection Helpers -----------------
def is_product_page(soup):
    """
    Heuristic check whether the page is a product detail page:
    - has <h3> product heading OR
    - has gallery-items div OR
    - has ul with li.bullet-text OR
    - has tab-parts id or tab-size id
    """
    if not soup:
        return False
    if soup.find("h3"):
        return True
    if soup.find("div", class_="gallery-items"):
        return True
    # ul with li.bullet-text
    for u in soup.find_all("ul"):
        if u.find("li", class_="bullet-text"):
            return True
    if soup.find(id="tab-parts") or soup.find(id="tab-size"):
        return True
    # fallback: presence of many product-like anchors or img names
    if soup.find("img", {"class": "pr-img"}):
        return True
    return False

def clean_size_text(raw):
    """
    Given a raw header like "Size - 7.5 feet" or "Size: 7.5 feet" or "Size 7.5 feet",
    return only the value part "7.5 feet".
    """
    if not raw:
        return ""
    txt = raw.strip()
    # Lowercase check for 'size' prefix, but preserve case of remainder
    # Remove common prefixes: "Size -", "Size:", "Size"
    # We'll find the first occurrence of digits/letters after 'Size'
    # Try simple replacements first
    for prefix in ["Size -", "Size:", "Size - ", "Size : ", "Size "]:
        if txt.startswith(prefix):
            return txt[len(prefix):].strip()
    # fallback remove leading 'Size' if present
    if txt.lower().startswith("size"):
        # remove the word 'size' and any separators
        remainder = txt[4:].strip(" :-")
        return remainder.strip()
    return txt

# ----------------- Parsing product page -----------------
def parse_product_page(html, product_url, category_name):
    soup = BeautifulSoup(html, "html.parser")
    out = {k: "" for k in COLUMNS}
    out["Category"] = category_name
    out["Product URL"] = product_url
    out["SKU"] = ""       # always empty per requirement
    out["Brand"] = "GALTECHCORP"

    # Product name
    h3 = soup.find("h3")
    if h3:
        out["Product Name"] = h3.get_text(strip=True)
    else:
        # fallback to title or first h1
        t = soup.find("title")
        if t:
            out["Product Name"] = t.get_text(strip=True)
        else:
            h1 = soup.find("h1")
            out["Product Name"] = h1.get_text(strip=True) if h1 else ""

    # Full Description (HTML) -> <ul> that contains li.bullet-text
    description_html = ""
    for u in soup.find_all("ul"):
        if u.find("li", class_="bullet-text"):
            description_html = str(u)
            break
    out["Full Description (HTML)"] = description_html

    # Attr_SizeText and Attr_SizeImage
    size_text = ""
    size_img = ""
    tab_size = soup.find(id="tab-size")
    if tab_size:
        h2 = tab_size.find("h2", class_="con-heading")
        if h2:
            raw = h2.get_text(strip=True)
            size_text = clean_size_text(raw)
        img = tab_size.find("img")
        if img and img.get("src"):
            size_img = abs_url(img.get("src"))
    else:
        # fallback search for any h2.con-heading that includes 'Size'
        for h2 in soup.find_all("h2", class_="con-heading"):
            if "size" in h2.get_text(strip=True).lower():
                raw = h2.get_text(strip=True)
                size_text = clean_size_text(raw)
                # try to find an image nearby
                next_img = h2.find_next("img")
                if next_img and next_img.get("src"):
                    size_img = abs_url(next_img.get("src"))
                break

    out["Attr_SizeText"] = size_text
    out["Attr_SizeImage"] = size_img

    # Part List (HTML): from id="tab-parts" but remove h2/p as requested
    part_html = ""
    tab_parts = soup.find(id="tab-parts")
    if tab_parts:
        copy = BeautifulSoup(str(tab_parts), "html.parser")
        for tag in copy.find_all(["h2", "p"]):
            tag.decompose()
        part_html = str(copy)
    else:
        # fallback: find tab-content with heading containing "Parts"
        for t in soup.find_all("div", class_="tab-content"):
            h2 = t.find("h2", class_="con-heading")
            if h2 and "part" in h2.get_text(strip=True).lower():
                copy = BeautifulSoup(str(t), "html.parser")
                for tag in copy.find_all(["h2", "p"]):
                    tag.decompose()
                part_html = str(copy)
                break
    out["Part List (HTML)"] = part_html

    # Gallery images up to MAX_IMAGES
    images = []
    gallery = soup.find("div", class_="gallery-items")
    if gallery:
        # prefer anchor hrefs (large images)
        for a in gallery.find_all("a"):
            href = a.get("href")
            if href:
                images.append(abs_url(href))
            # if we already have enough
            if len(images) >= MAX_IMAGES:
                break
        if not images:
            # fallback to <img> tags
            for img in gallery.find_all("img"):
                src = img.get("src")
                if src:
                    images.append(abs_url(src))
                if len(images) >= MAX_IMAGES:
                    break
    else:
        # another fallback: any images with class pr-img in page
        for img in soup.find_all("img", class_="pr-img"):
            src = img.get("src")
            if src:
                images.append(abs_url(src))
            if len(images) >= MAX_IMAGES:
                break

    for i in range(MAX_IMAGES):
        out[f"Image {i+1}"] = images[i] if i < len(images) else ""

    return out

# ----------------- MAIN -----------------
def main():
    global runtime_driver, rows
    driver = new_driver()
    runtime_driver = driver

    try:
        print("🔍 Loading homepage...")
        driver.get(BASE_URL)
        time.sleep(PAGE_SLEEP)

        html = driver.page_source
        soup = BeautifulSoup(html, "html.parser")
        nav_holder = soup.find("div", class_="nav-holder")
        if not nav_holder:
            print("⚠️ nav-holder not found. Aborting.")
            save_now(reason="nav_holder_not_found")
            driver.quit()
            return

        # collect categories from CATEGORY_START -> CATEGORY_END
        collect = False
        categories = []
        for li in nav_holder.select("nav > ul > li"):
            a = li.find("a", recursive=False)
            if not a:
                continue
            cat_name = a.get_text(strip=True)
            if cat_name == CATEGORY_START:
                collect = True
            if collect:
                # gather sub-product links from ul if present
                sub_links = []
                ul = li.find("ul")
                if ul:
                    for sub_a in ul.find_all("a", recursive=True):
                        name = sub_a.get_text(strip=True)
                        href = sub_a.get("href")
                        if href:
                            sub_links.append((name, abs_url(href)))
                categories.append((cat_name, sub_links))
            if cat_name == CATEGORY_END:
                break

        # process categories
        for cat_name, products in categories:
            if not products:
                print(f"ℹ️ Category '{cat_name}' has no product links, skipping.")
                continue

            total = len(products)
            print(f"\n📁 Category '{cat_name}' -> {total} product(s) found.")

            for idx, (nav_prod_name, prod_url) in enumerate(products, start=1):
                try:
                    print(f"Scraping {idx}/{total} products from {cat_name} -> {nav_prod_name}")

                    # Load product-ish page
                    driver.get(prod_url)
                    time.sleep(PAGE_SLEEP)

                    page_html = driver.page_source
                    page_soup = BeautifulSoup(page_html, "html.parser")

                    # Skip if not a product detail page (heuristic)
                    if not is_product_page(page_soup):
                        print(f"  ⚠️ Skipping non-product (landing) page: {prod_url}")
                        continue

                    record = parse_product_page(page_html, prod_url, cat_name)

                    # fallback: if Product Name empty, use nav text
                    if not record.get("Product Name"):
                        record["Product Name"] = nav_prod_name

                    # append row
                    rows.append([record.get(c, "") for c in COLUMNS])

                    # autosave after every product
                    save_now(reason=f"after product {idx}/{total} in {cat_name}")

                except Exception as e:
                    # safe continue & autosave partial progress
                    print(f"  ❌ Error scraping {prod_url}: {e}")
                    save_now(reason=f"error_on_{prod_url}")
                    continue

        # final save
        save_now(reason="final_save")
        print("\n🎉 Extraction finished.")

    except Exception as e:
        print("❌ Fatal error in main:", e)
        save_now(reason=f"fatal_{e}")
    finally:
        try:
            driver.quit()
        except Exception:
            pass

if __name__ == "__main__":
    main()