import time
import re
import signal
import sys
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

# ------------------ CONFIG ------------------
BASE_URL = "https://www.seasonalliving.com"
SAVE_FILE = "seasonalliving_products_partial.xlsx"
FINAL_FILE = "seasonalliving_products_final.xlsx"
SAVE_INTERVAL = 50

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=chrome_options)
all_rows = []

# ------------------ SAVE & INTERRUPT ------------------
def save_progress(filename=SAVE_FILE):
    if all_rows:
        df = pd.DataFrame(all_rows)
        df.to_excel(filename, index=False)
        print(f"\n💾 Progress saved: {len(df)} rows → {filename}")

def signal_handler(sig, frame):
    print("\n⚠️ Script interrupted. Saving progress before exit...")
    save_progress()
    driver.quit()
    sys.exit(0)

signal.signal(signal.SIGINT, signal_handler)

# ------------------ UTILITIES ------------------
def safe_text(element):
    return element.text.strip() if element else ""

# def scrape_product(product_url):
#     try:
#         driver.get(product_url)
#         time.sleep(1)

#         # Product Name
#         try:
#             h1_element = driver.find_element(By.CSS_SELECTOR, "h1.product-title")
#             product_name = driver.execute_script("""
#                 var el = arguments[0];
#                 var childNodes = el.childNodes;
#                 var text = '';
#                 for (var i=0; i<childNodes.length; i++){
#                     if(childNodes[i].nodeType === Node.TEXT_NODE){
#                         text += childNodes[i].textContent.trim();
#                     }
#                 }
#                 return text;
#             """, h1_element)
#         except:
#             product_name = ""

#         # SKU
#         try:
#             sku = driver.find_element(By.CSS_SELECTOR, "table.productSpecs th").text.strip().replace(":", "")
#         except:
#             sku = ""

#         # Brand
#         brand = "Seasonal Living"

#         # Category Path
#         try:
#             breadcrumb_elements = driver.find_elements(By.CSS_SELECTOR, "nav.woocommerce-breadcrumb a")
#             category_path = " / ".join([el.text.strip() for el in breadcrumb_elements] + [product_name])
#         except:
#             category_path = product_name

#         # Short Description
#         short_desc = ""
#         try:
#             desc_elements = driver.find_elements(By.CSS_SELECTOR, "div.sl-accordian p")
#             short_desc = " ".join([el.text.strip() for el in desc_elements])
#         except:
#             pass

#         # Full HTML Description
#         try:
#             full_html_element = driver.find_element(By.CSS_SELECTOR, "div.sl-accordian")
#             full_html = full_html_element.get_attribute("outerHTML")
#         except:
#             full_html = ""

#         # Tear Sheet URL
#         try:
#             tear_element = driver.find_element(By.CSS_SELECTOR, "a.print-tear-sheet")
#             tear_url = tear_element.get_attribute("href")
#             if tear_url and not tear_url.startswith("http"):
#                 tear_url = BASE_URL + tear_url
#         except:
#             tear_url = ""

#         # Images (up to 4)
#         images = []
#         try:
#             img_elements = driver.find_elements(By.CSS_SELECTOR, "ol.flex-control-thumbs img")
#             for img in img_elements[:4]:
#                 images.append(img.get_attribute("src"))
#         except:
#             pass
#         while len(images) < 4:
#             images.append("")

#         return {
#             "Category Path": category_path,
#             "Product URL": product_url,
#             "Product Name": product_name,
#             "SKU": sku,
#             "Brand": brand,
#             "Short Description": short_desc,
#             "Full Description HTML": full_html,
#             "Tear Sheet URL": tear_url,
#             "Image1": images[0],
#             "Image2": images[1],
#             "Image3": images[2],
#             "Image4": images[3],
#         }

#     except Exception as e:
#         print(f"❌ Error scraping product {product_url}: {e}")
#         return None

def _clean_attr_text(raw):
    """
    Clean innerHTML/innerText from the product spec TD:
    - convert &nbsp; / \xa0 to space
    - replace <br> with ' | '
    - strip remaining tags
    - collapse multiple whitespace
    """
    if not raw:
        return ""
    # convert to str (in case it's None)
    s = str(raw)
    # replace HTML non-breaking space entities
    s = s.replace("&nbsp;", " ").replace("\xa0", " ")
    # normalize <br> to a separator (handle common variants)
    s = re.sub(r"(?i)<br\s*/?>", " | ", s)
    # remove any remaining tags
    s = re.sub(r"<[^>]+>", " ", s)
    # collapse whitespace
    s = re.sub(r"\s+", " ", s)
    return s.strip()

# ------------------ SCRAPE PRODUCT (replacement) ------------------
def scrape_product(product_url):
    try:
        driver.get(product_url)
        time.sleep(1.5)

        # ---------------- CATEGORY PATH ----------------
        try:
            breadcrumb = driver.find_element(By.CSS_SELECTOR, "nav.woocommerce-breadcrumb")
            category_path = driver.execute_script("""
                var el = arguments[0];
                var text = el.innerText || '';
                return text.replace(/\\s*›\\s*/g, ' / ')
                           .replace(/\\s*·\\s*/g, ' / ')
                           .replace(/\\s*\\/\\s*/g, ' / ')
                           .replace(/\\s+/g, ' ')
                           .trim();
            """, breadcrumb)
        except:
            category_path = ""

        # ---------------- PRODUCT NAME ----------------
        try:
            h1 = driver.find_element(By.CSS_SELECTOR, "h1.product-title")
            product_name = driver.execute_script("""
                var el = arguments[0];
                var text = '';
                el.childNodes.forEach(n => {
                    if (n.nodeType === Node.TEXT_NODE) text += n.textContent.trim();
                });
                return text.trim();
            """, h1)
        except:
            product_name = ""

        # ---------------- BRAND ----------------
        brand = "Seasonal Living"

        # ---------------- SKU ----------------
        sku = ""
        try:
            sku_element = driver.find_element(By.CSS_SELECTOR, "div.sl-accordian-container div#ProductInformation table.productSpecs th")
            sku = sku_element.text.strip().rstrip(":")
        except:
            pass

        # ---------------- PRODUCT INFORMATION ACCORDION (DESCRIPTION) ----------------
        short_desc = ""
        try:
            # Locate Product Information accordion
            info_title = driver.find_element(By.XPATH, "//div[contains(@class,'sl-accordian-title')]/h2[contains(.,'Product Information')]")
            info_container = info_title.find_element(By.XPATH, "./ancestor::div[contains(@class,'sl-accordian-container')]")
            info_content = info_container.find_element(By.XPATH, ".//div[contains(@class,'sl-accordian-content')]")

            # Check if closed → open it
            style = info_content.get_attribute("style")
            if "display: none" in style or "max-height: 0" in style:
                driver.execute_script("arguments[0].scrollIntoView(true);", info_title)
                driver.execute_script("arguments[0].click();", info_title)
                time.sleep(1.2)

            short_desc = info_content.text.strip()
        except:
            # fallback: simple short description area
            try:
                short_desc = driver.find_element(By.CSS_SELECTOR, "div.woocommerce-product-details__short-description").text.strip()
            except:
                short_desc = ""

        # ---------------- FULL DESCRIPTION HTML ----------------
        full_html = ""
        try:
            accordian = driver.find_element(By.CSS_SELECTOR, "div.sl-accordian")
            full_html = accordian.get_attribute("outerHTML")
        except:
            pass

        # ---------------- DIMENSIONS ACCORDION ----------------
        attr_dimension = ""
        attr_weight = ""
        attr_packing_info = ""
        try:
            dim_title = driver.find_element(By.XPATH, "//div[contains(@class,'sl-accordian-title')]/h2[contains(.,'Dimensions')]")
            dim_container = dim_title.find_element(By.XPATH, "./ancestor::div[contains(@class,'sl-accordian-container')]")
            dim_content = dim_container.find_element(By.XPATH, ".//div[contains(@class,'sl-accordian-content')]")

            # If closed, click to open
            style = dim_content.get_attribute("style")
            if "display: none" in style or "max-height: 0" in style:
                driver.execute_script("arguments[0].scrollIntoView(true);", dim_title)
                driver.execute_script("arguments[0].click();", dim_title)
                time.sleep(1.2)

            # Extract table data
            rows = dim_content.find_elements(By.CSS_SELECTOR, "table.productSpecs tr")
            for row in rows:
                th = row.find_element(By.TAG_NAME, "th").text.strip()
                td_html = row.find_element(By.TAG_NAME, "td").get_attribute("innerHTML")
                clean_val = _clean_attr_text(td_html)
                if "dimension" in th.lower():
                    attr_dimension = clean_val
                elif "weight" in th.lower():
                    attr_weight = clean_val
                elif "pack" in th.lower():
                    attr_packing_info = clean_val
        except Exception as e:
            print(f"⚠️ Dimensions accordion issue on {product_url}: {e}")

        # ---------------- TEAR SHEET ----------------
        tear_url = ""
        try:
            tear_element = driver.find_element(By.CSS_SELECTOR, "a.print-tear-sheet")
            tear_url = tear_element.get_attribute("href")
            if tear_url and not tear_url.startswith("http"):
                tear_url = BASE_URL + tear_url
        except:
            pass

        # ---------------- IMAGES ----------------
        images = []
        try:
            main_gallery = driver.find_elements(By.CSS_SELECTOR, "div.woocommerce-product-gallery__image img")
            for img in main_gallery:
                src = img.get_attribute("src")
                if src and src not in images:
                    images.append(src)
        except:
            pass

        while len(images) < 4:
            images.append("")

        # ---------------- RETURN FINAL DATA ----------------
        return {
            "Category Path": category_path,
            "Product URL": product_url,
            "Product Name": product_name,
            "SKU": sku,
            "Brand": brand,
            "Short Description": short_desc,
            "Full Description HTML": full_html,
            "Attr_Dimension": attr_dimension,
            "Attr_Weight": attr_weight,
            "Attr_PackingInfo": attr_packing_info,
            "Tear Sheet URL": tear_url,
            "Image1": images[0],
            "Image2": images[1],
            "Image3": images[2],
            "Image4": images[3],
        }

    except Exception as e:
        print(f"❌ Error scraping product {product_url}: {e}")
        return None

# ------------------ SCRAPE CATEGORY ------------------
def scrape_category(category_name, category_url):
    try:
        driver.get(category_url)
        time.sleep(2)
        print(f"\n🗂️ Scraping category: {category_name}")

        # Infinite scroll: scroll until no new products load
        last_height = driver.execute_script("return document.body.scrollHeight")
        while True:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height

        # Collect all product links
        product_elements = driver.find_elements(By.CSS_SELECTOR, "a.woocommerce-LoopProduct-link")
        product_links = [el.get_attribute("href") for el in product_elements]
        print(f"🔹 Total products found in {category_name}: {len(product_links)}")

        # Scrape each product
        total_scraped = 0
        for idx, product_url in enumerate(product_links, start=1):
            row = scrape_product(product_url)
            if row:
                all_rows.append(row)
                total_scraped += 1
            print(f"✅ {category_name}: Scraped {idx}/{len(product_links)} (Total {total_scraped})")

            if len(all_rows) % SAVE_INTERVAL == 0:
                save_progress()

        print(f"🏁 Finished category {category_name} — Total scraped: {total_scraped}")

    except Exception as e:
        print(f"⚠️ Error in category {category_name}: {e}")

# ------------------ MAIN EXECUTION ------------------
if __name__ == "__main__":
    try:
        categories = {
            "Signature Collections": "https://www.seasonalliving.com/collection/signature-collections/",
            "Eterna Collection": "https://www.seasonalliving.com/collection/eterna-collection/",
            "Vaterra Collection": "https://www.seasonalliving.com/collection/vaterra-collection/",
            "Patinero Collection": "https://www.seasonalliving.com/collection/patinero-collection/",
            "Serenique Collection": "https://www.seasonalliving.com/collection/serenique-collection/",
            "Companion Collections": "https://www.seasonalliving.com/collection/companion-collections/",
            "Ceramics Companion": "https://www.seasonalliving.com/collection/ceramics-companion-collection/",
            "Lightweight Concrete Companion": "https://www.seasonalliving.com/collection/lightweight-concrete-companion-collection/",
            "Durafiber Companion": "https://www.seasonalliving.com/collection/durafiber-companion-collection/",
            "Upholstery Companion": "https://www.seasonalliving.com/collection/upholstery-companion-collection/",
            "Provenance Ceramics": "https://www.seasonalliving.com/collection/provenance-ceramics-collection/",
            "Perpetual Collection": "https://www.seasonalliving.com/collection/perpetual-collection/",
            "Explorer": "https://www.seasonalliving.com/collection/explorer/",
            "Archipelago Collection": "https://www.seasonalliving.com/collection/archipelago-collection/",
            "Ceramics Material": "https://www.seasonalliving.com/material/ceramics/",
            "Lightweight Concrete Material": "https://www.seasonalliving.com/material/lightweight-concrete-stone/",
            "Durafiber Material": "https://www.seasonalliving.com/material/durafiber-fiber-reinforced-polymer-frp/",
            "Outdoor Fabric Upholstery": "https://www.seasonalliving.com/material/outdoor-fabric-upholstery/",
            "Metal Furniture": "https://www.seasonalliving.com/material/metal-furniture/",
            "Wood": "https://www.seasonalliving.com/material/wood/",
            "Woven": "https://www.seasonalliving.com/material/woven/",
            "Upholstery Sofas Sectionals": "https://www.seasonalliving.com/product-category/outdoor-indoor-furniture/upholstery-sofas-sectionals/",
            "Dining": "https://www.seasonalliving.com/product-category/outdoor-indoor-furniture/dining/",
            "Dining Chairs": "https://www.seasonalliving.com/product-category/outdoor-indoor-furniture/dining/dining-chairs/",
            "Dining Tables": "https://www.seasonalliving.com/product-category/outdoor-indoor-furniture/dining/dining-tables/",
            "Tables": "https://www.seasonalliving.com/product-category/outdoor-indoor-furniture/tables/",
            "Coffee Tables": "https://www.seasonalliving.com/product-category/outdoor-indoor-furniture/tables/coffee-tables/",
            "Side Tables": "https://www.seasonalliving.com/product-category/outdoor-indoor-furniture/tables/side-tables/",
            "Seating Chairs Lounge": "https://www.seasonalliving.com/product-category/outdoor-indoor-furniture/seating-chairs-lounge/",
            "Bar Counter Chairs": "https://www.seasonalliving.com/product-category/outdoor-indoor-furniture/seating-chairs-lounge/bar-counter-chairs/",
            "Benches": "https://www.seasonalliving.com/product-category/outdoor-indoor-furniture/seating-chairs-lounge/benches/",
            "Chaise": "https://www.seasonalliving.com/product-category/outdoor-indoor-furniture/seating-chairs-lounge/chaise/",
            "Sofas": "https://www.seasonalliving.com/product-category/outdoor-indoor-furniture/seating-chairs-lounge/sofas/",
            "Sectionals": "https://www.seasonalliving.com/product-category/outdoor-indoor-furniture/seating-chairs-lounge/sectionals/",
            "Upholstery Ottomans": "https://www.seasonalliving.com/product-category/outdoor-indoor-furniture/upholstery-sofas-sectionals/ottomans/",
            "Lounge": "https://www.seasonalliving.com/product-category/outdoor-indoor-furniture/seating-chairs-lounge/lounge/",
            "Stools": "https://www.seasonalliving.com/product-category/outdoor-indoor-furniture/seating-chairs-lounge/stools/",
            "Planters": "https://www.seasonalliving.com/product-category/planters/",
            "Shop": "https://www.seasonalliving.com/shop/",
            "Red": "https://www.seasonalliving.com/shop/?color_families%5B%5D=red",
            "Orange": "https://www.seasonalliving.com/shop/?color_families%5B%5D=orange",
            "Yellow": "https://www.seasonalliving.com/shop/?color_families%5B%5D=yellow",
            "Green": "https://www.seasonalliving.com/shop/?color_families%5B%5D=green",
            "Blue": "https://www.seasonalliving.com/shop/?color_families%5B%5D=blue",
            "Purple": "https://www.seasonalliving.com/shop/?color_families%5B%5D=purple",
            "White": "https://www.seasonalliving.com/shop/?color_families%5B%5D=white",
            "Gray": "https://www.seasonalliving.com/shop/?color_families%5B%5D=gray",
            "Black": "https://www.seasonalliving.com/shop/?color_families%5B%5D=black",
            "Brown": "https://www.seasonalliving.com/shop/?color_families%5B%5D=brown",
            "Tan": "https://www.seasonalliving.com/shop/?color_families%5B%5D=tan",
            "Metallic": "https://www.seasonalliving.com/shop/?color_families%5B%5D=metallic",
            "New Products": "https://www.seasonalliving.com/product-tag/new-products/",
            "Hospitality": "https://www.seasonalliving.com/product-tag/hospitality/",
            "Upholstered": "https://www.seasonalliving.com/product-category/upholstered/",
        }

        for name, url in categories.items():
            scrape_category(name, url)

        save_progress(FINAL_FILE)
        print("\n🎉 Scraping completed successfully!")

    except Exception as e:
        print(f"💥 Fatal error: {e}")
        save_progress()
    finally:
        driver.quit()
        print("👋 Browser closed.")
