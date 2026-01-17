import time
import re
import pandas as pd
import os
import traceback
import random
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ---- CONFIGURE CHROME ----
chrome_options = Options()
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--window-size=1920,1080")
# mimic a normal browser UA
chrome_options.add_argument(
    "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36"
)
# chrome_options.add_argument("--headless")  # uncomment for headless mode if desired

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
wait = WebDriverWait(driver, 18)

# ---- CATEGORY URLS ----
category_urls = [
    "https://www.jensenoutdoor.com/collections/tempo/",
    "https://www.jensenoutdoor.com/collections/sorrento/",
    "https://www.jensenoutdoor.com/collections/savannah/",
    "http://www.jensenoutdoor.com/collections/foundations/",
    "http://www.jensenoutdoor.com/collections/glow/",
    "http://www.jensenoutdoor.com/collections/mix/",
    "http://www.jensenoutdoor.com/collections/sky/",
    "https://www.jensenoutdoor.com/collections/opal/",
    "https://www.jensenoutdoor.com/collections/classic-ipe/",
    "https://www.jensenoutdoor.com/collections/coral/",
    "https://www.jensenoutdoor.com/collections/richmond/",
    "https://www.jensenoutdoor.com/collections/laguna/",
    "https://www.jensenoutdoor.com/collections/forte/",
    "https://www.jensenoutdoor.com/collections/topaz/",
    "https://www.jensenoutdoor.com/collections/vintage/",
    "https://www.jensenoutdoor.com/collections/nest/",
    "https://www.jensenoutdoor.com/collections/jett/",
    "https://www.jensenoutdoor.com/collections/unicon/",
    "https://www.jensenoutdoor.com/collections/harmony/",
    "https://www.jensenoutdoor.com/collections/dana/",
    "https://www.jensenoutdoor.com/collections/inception/",
    "https://www.jensenoutdoor.com/collections/innova/",
    "https://www.jensenoutdoor.com/collections/plume-pillows/",
    "https://www.jensenoutdoor.com/collections/velo/",
    "https://www.jensenoutdoor.com/outdoor-furniture/new/",
    "https://www.jensenoutdoor.com/outdoor-furniture/best-sellers/",
    "https://www.jensenoutdoor.com/dining/",
    "https://www.jensenoutdoor.com/dining/tables/",
    "https://www.jensenoutdoor.com/dining/chairs/",
    "https://www.jensenoutdoor.com/dining/bar/",
    "https://www.jensenoutdoor.com/lounging/",
    "https://www.jensenoutdoor.com/lounging/accessory-tables/",
    "https://www.jensenoutdoor.com/lounging/adirondack/",
    "https://www.jensenoutdoor.com/lounging/chaise-lounges/",
    "https://www.jensenoutdoor.com/lounging/deep-seating/",
    "https://www.jensenoutdoor.com/lounging/deep-seating/lounge-chairs/",
    "https://www.jensenoutdoor.com/lounging/deep-seating/sofas-loveseats/",
    "https://www.jensenoutdoor.com/accessories/",
    "https://www.jensenoutdoor.com/accessories/cushions/",
    "https://www.jensenoutdoor.com/accessories/lazy-susan/",
    "https://www.jensenoutdoor.com/accessories/parts/",
    "https://www.jensenoutdoor.com/accessories/serving-tray/",
    "https://www.jensenoutdoor.com/accessories/table-extensions/",
    "https://www.jensenoutdoor.com/lounging/deep-seating/sectionals/",
    "https://www.jensenoutdoor.com/dining/tables/extension-tables/",
    "https://www.jensenoutdoor.com/lounging/garden-benches/"
]

# ---- DATA STORAGE ----
products_data = []
SAVE_FILE = "jensen_products.xlsx"
SAVE_INTERVAL = 50

def save_data():
    if products_data:
        df = pd.DataFrame(products_data)
        df.to_excel(SAVE_FILE, index=False)
        print(f"💾 Auto-saved {len(products_data)} products to {SAVE_FILE}")

def scroll_to_bottom(max_scrolls=6, sleep=1.5):
    last_height = driver.execute_script("return document.body.scrollHeight")
    scrolls = 0
    while scrolls < max_scrolls:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(sleep + random.random()*0.8)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height
        scrolls += 1

def safe_text_find(context, by, selector, default=""):
    try:
        elm = context.find_element(by, selector)
        return elm.text.strip()
    except:
        return default

def safe_attr_find(context, by, selector, attr, default=""):
    try:
        elm = context.find_element(by, selector)
        return elm.get_attribute(attr) or default
    except:
        return default

def clean_image_url(url):
    """Remove WordPress size suffixes like -80x60 before the extension to get full-size image."""
    if not url:
        return ""
    m = re.search(r"^(.+?)(-\d+x\d+)(\.[a-zA-Z0-9]+)$", url)
    if m:
        return m.group(1) + m.group(3)
    return url

def extract_images_from_thumbnails():
    """Extract up to 4 images from div.product-thumbnails anchors (style background-image)."""
    images = []
    try:
        thumb_anchors = driver.find_elements(By.CSS_SELECTOR, "div.product-thumbnails div.product-mini-image a")
        for a in thumb_anchors:
            if len(images) >= 4:
                break
            style = a.get_attribute("style") or ""
            m = re.search(r"url\((['\"]?)(.*?)\1\)", style)
            if m:
                src = m.group(2)
                src = clean_image_url(src)
                if src and src not in images:
                    images.append(src)
    except:
        pass

    # fallback: check for <img> tags in product gallery
    if len(images) < 4:
        try:
            img_tags = driver.find_elements(By.CSS_SELECTOR, "div.product-gallery img, div.woocommerce-product-gallery__image img, img.wp-post-image")
            for img in img_tags:
                if len(images) >= 4:
                    break
                src = img.get_attribute("src") or img.get_attribute("data-src") or ""
                if src:
                    images.append(clean_image_url(src))
        except:
            pass

    # ensure length 4
    while len(images) < 4:
        images.append("")
    return images[:4]

def get_product_links_from_category(category_url):
    """Return ordered unique product links found in the category's product list (li elements)."""
    driver.get(category_url)
    # wait for the product list or a query block to appear
    try:
        wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "ul.wp-block-post-template, div.wp-block-query, ul.columns-3")
            )
        )
    except:
        # still continue (some pages maybe slightly different)
        pass

    # ensure lazy load triggers
    scroll_to_bottom()

    # try several candidate container selectors then collect li children
    candidate_selectors = [
        "ul.wp-block-post-template",
        "div.wp-block-query",
        "ul.columns-3",
        "ul.wp-block-post-template.is-layout-grid",
    ]

    product_li_elements = []
    for sel in candidate_selectors:
        try:
            container = driver.find_element(By.CSS_SELECTOR, sel)
            # get relevant li children
            # match common li classes used on site: li.post.product, li.product-item, li.product
            lis = container.find_elements(By.CSS_SELECTOR, "li.post.product, li.product-item, li.product")
            if len(lis) == 0:
                # try any li
                lis = container.find_elements(By.TAG_NAME, "li")
            if lis:
                product_li_elements = lis
                break
        except:
            continue

    # final fallback: find matching lis across whole page
    if not product_li_elements:
        product_li_elements = driver.find_elements(By.CSS_SELECTOR, "ul li.post.product, ul li.product-item, ul li.product")

    # extract link from each li using the title anchor first, then figure anchor, then first href containing domain
    links = []
    seen = set()
    for li in product_li_elements:
        link = ""
        # 1) title anchor
        try:
            a = li.find_element(By.CSS_SELECTOR, "h2.wp-block-post-title a, h2.wp-block-post-title a")
            link = a.get_attribute("href") or ""
        except:
            link = ""
        # 2) figure anchor fallback
        if not link:
            try:
                a = li.find_element(By.CSS_SELECTOR, "figure a")
                link = a.get_attribute("href") or ""
            except:
                link = ""
        # 3) any anchor inside li whose href looks like product
        if not link:
            try:
                anchors = li.find_elements(By.TAG_NAME, "a")
                for a in anchors:
                    href = a.get_attribute("href") or ""
                    if href and "jensenoutdoor.com" in href and href not in ("#", ""):
                        link = href
                        break
            except:
                pass

        # normalize and deduplicate
        if link:
            link = link.strip()
            if link.endswith("/#") or link.endswith("#"):
                link = link.split("#")[0]
            if link not in seen:
                seen.add(link)
                links.append(link)

    return links

# ---- MAIN SCRAPER ----
# try:
#     # resume if file exists
#     if os.path.exists(SAVE_FILE):
#         old_df = pd.read_excel(SAVE_FILE)
#         products_data = old_df.to_dict("records")
#         print(f"🔁 Resuming from previous progress ({len(products_data)} records).")

#     for category_url in category_urls:
#         print("\n" + "="*80)
#         print(f"📂 Processing category: {category_url}")
#         product_links = get_product_links_from_category(category_url)
#         total_products = len(product_links)
#         print(f"Found {total_products} product links in category.")

#         for idx, product_link in enumerate(product_links, start=1):
#             # small random sleep to mimic human behaviour
#             time.sleep(1.0 + random.random()*1.5)

#             # Robust navigation with a couple retries if redirect to category happens
#             success = False
#             for attempt in range(3):
#                 try:
#                     driver.get(product_link)
#                     # wait for a product H1 title to appear
#                     wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "h1.wp-block-post-title, h1.product_title, h1")))
#                     current_url = driver.current_url
#                     # detect if server redirected us back to a collections/category page
#                     if "/collections/" in current_url and current_url.rstrip("/") == category_url.rstrip("/"):
#                         # redirected to the same category page -> retry after a pause
#                         print(f"⚠️ Redirected back to category for {product_link} (attempt {attempt+1}). Retrying...")
#                         time.sleep(2 + random.random()*2)
#                         continue

#                     # Extract product details
#                     name = ""
#                     try:
#                         name = driver.find_element(By.CSS_SELECTOR, "h1.wp-block-post-title, h1.product_title, h1").text.strip()
#                     except:
#                         name = ""

#                     sku = ""
#                     try:
#                         sku = driver.find_element(By.CSS_SELECTOR, "span.sku, .product-sku").text.strip()
#                     except:
#                         sku = ""

#                     price = ""
#                     try:
#                         price = driver.find_element(By.CSS_SELECTOR, "span.woocommerce-Price-amount, .price, .ProductMeta__Price").text.strip()
#                     except:
#                         price = ""

#                     breadcrumb = ""
#                     try:
#                         breadcrumb = driver.find_element(By.CSS_SELECTOR, "nav.woocommerce-breadcrumb, nav.breadcrumb").text.strip()
#                     except:
#                         breadcrumb = ""

#                     # descriptions
#                     short_desc = ""
#                     try:
#                         short_desc = driver.find_element(By.CSS_SELECTOR, "div.block-accordion-item-content-inner, .product-short-description, .ProductMeta__Description").get_attribute("innerHTML") or ""
#                     except:
#                         short_desc = ""

#                     full_desc = ""
#                     try:
#                         full_desc = driver.find_element(By.CSS_SELECTOR, "div.block-accordion, .product-description, .Product__Accordion").get_attribute("outerHTML") or ""
#                     except:
#                         full_desc = ""

#                     # images (from product-thumbnails style backgrounds first)
#                     image1, image2, image3, image4 = extract_images_from_thumbnails()

#                     products_data.append({
#                         "Category": breadcrumb,
#                         "Product URL": current_url,
#                         "Product Name": name,
#                         "SKU": sku,
#                         "Price": price,
#                         "Brand": "Jensen Outdoor",
#                         "Short Description": short_desc,
#                         "Full Description HTML": full_desc,
#                         "Image1": image1,
#                         "Image2": image2,
#                         "Image3": image3,
#                         "Image4": image4,
#                     })

#                     print(f"✅ ({idx}/{total_products}) {name} -- {current_url}")
#                     success = True

#                     # auto-save
#                     if len(products_data) % SAVE_INTERVAL == 0:
#                         save_data()

#                     break  # break retry loop

#                 except Exception as e:
#                     print(f"❌ Error scraping product {idx} ({product_link}) on attempt {attempt+1}: {e}")
#                     traceback.print_exc()
#                     time.sleep(2 + random.random()*2)
#                     continue

#             if not success:
#                 print(f"⚠️ Skipped product after retries: {product_link}")
#                 # continue to next product

#         print(f"✅ Completed category: {category_url}")

# except KeyboardInterrupt:
#     print("🛑 Script manually stopped. Saving progress...")

# except Exception as e:
#     print(f"💥 Unexpected crash: {e}")
#     traceback.print_exc()

# finally:
#     save_data()
#     driver.quit()
#     print("🏁 Scraping completed.")


try:
    if os.path.exists(SAVE_FILE):
        old_df = pd.read_excel(SAVE_FILE)
        products_data = old_df.to_dict("records")
        print(f"🔁 Resuming from previous progress ({len(products_data)} records).")

    for category_url in category_urls:
        print("\n" + "="*80)
        print(f"📂 Processing category: {category_url}")
        product_links = get_product_links_from_category(category_url)
        total_products = len(product_links)
        print(f"Found {total_products} product links in category.")

        for idx, product_link in enumerate(product_links, start=1):
            time.sleep(1.0 + random.random()*1.5)
            success = False

            for attempt in range(3):
                try:
                    driver.get(product_link)
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "h1.wp-block-post-title, h1.product_title, h1")))
                    current_url = driver.current_url
                    if "/collections/" in current_url and current_url.rstrip("/") == category_url.rstrip("/"):
                        print(f"⚠️ Redirected back to category for {product_link} (attempt {attempt+1}). Retrying...")
                        time.sleep(2 + random.random()*2)
                        continue

                    name = ""
                    try:
                        name = driver.find_element(By.CSS_SELECTOR, "h1.wp-block-post-title, h1.product_title, h1").text.strip()
                    except:
                        pass

                    sku = ""
                    try:
                        sku = driver.find_element(By.CSS_SELECTOR, "span.sku, .product-sku").text.strip()
                    except:
                        pass

                    price = ""
                    try:
                        price = driver.find_element(By.CSS_SELECTOR, "span.woocommerce-Price-amount, .price, .ProductMeta__Price").text.strip()
                    except:
                        pass

                    breadcrumb = ""
                    try:
                        breadcrumb = driver.find_element(By.CSS_SELECTOR, "nav.woocommerce-breadcrumb, nav.breadcrumb").text.strip()
                    except:
                        pass

                    short_desc = ""
                    try:
                        short_desc = driver.find_element(By.CSS_SELECTOR, "div.block-accordion-item-content-inner, .product-short-description, .ProductMeta__Description").get_attribute("innerHTML") or ""
                    except:
                        pass

                    full_desc = ""
                    try:
                        full_desc = driver.find_element(By.CSS_SELECTOR, "div.block-accordion, .product-description, .Product__Accordion").get_attribute("outerHTML") or ""
                    except:
                        pass

                    # ----------------------------------------
                    # ATTRIBUTE EXTRACTION (auto-expand + extract)
                    # ----------------------------------------
                    from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException

                    attr_weight = ""
                    attr_length = ""
                    attr_width = ""
                    attr_height = ""
                    attr_seat_height = ""
                    attr_arm_height = ""

                    try:
                        # 1️⃣ Locate the "Dimensions" accordion header
                        accordion_header = driver.find_element(
                            By.XPATH,
                            "//div[contains(@class,'block-accordion-item-header')][.//strong[contains(text(),'Dimensions')]]"
                        )

                        # 2️⃣ Click to expand (if not already expanded)
                        try:
                            driver.execute_script("arguments[0].scrollIntoView(true);", accordion_header)
                            time.sleep(0.5)
                            accordion_header.click()
                            time.sleep(1)
                        except ElementNotInteractableException:
                            pass  # Sometimes it's already expanded

                        # 3️⃣ Now extract data from the open content
                        li_elements = driver.find_elements(
                            By.XPATH,
                            "//div[contains(@class,'block-accordion-item-content-inner')]//li[.//strong and .//span]"
                        )

                        for li in li_elements:
                            try:
                                label = li.find_element(By.TAG_NAME, "strong").text.strip().replace(":", "").lower()
                                value = li.find_element(By.TAG_NAME, "span").text.strip()

                                if label == "weight":
                                    attr_weight = value
                                elif label == "length":
                                    attr_length = value
                                elif label == "width":
                                    attr_width = value
                                elif label == "height":
                                    attr_height = value
                                elif label == "seat height":
                                    attr_seat_height = value
                                elif label == "arm height":
                                    attr_arm_height = value

                            except Exception:
                                continue

                    except NoSuchElementException:
                        print(f"⚠️ No Dimensions section found for {current_url}")
                    except Exception as e:
                        print(f"⚠️ Dimension extraction error for {current_url}: {e}")

                    image1, image2, image3, image4 = extract_images_from_thumbnails()

                    products_data.append({
                        "Category": breadcrumb,
                        "Product URL": current_url,
                        "Product Name": name,
                        "SKU": sku,
                        "Price": price,
                        "Brand": "Jensen Outdoor",
                        "Short Description": short_desc,
                        "Full Description HTML": full_desc,
                        "Attr_Weight": attr_weight,
                        "Attr_Length": attr_length,
                        "Attr_Width": attr_width,
                        "Attr_Height": attr_height,
                        "Attr_SeatHeight": attr_seat_height,
                        "Attr_ArmHeight": attr_arm_height,
                        "Image1": image1,
                        "Image2": image2,
                        "Image3": image3,
                        "Image4": image4,
                    })

                    print(f"✅ ({idx}/{total_products}) {name} -- {current_url}")
                    success = True

                    if len(products_data) % SAVE_INTERVAL == 0:
                        save_data()

                    break

                except Exception as e:
                    print(f"❌ Error scraping product {idx} ({product_link}) on attempt {attempt+1}: {e}")
                    traceback.print_exc()
                    time.sleep(2 + random.random()*2)
                    continue

            if not success:
                print(f"⚠️ Skipped product after retries: {product_link}")

        print(f"✅ Completed category: {category_url}")

except KeyboardInterrupt:
    print("🛑 Script manually stopped. Saving progress...")
except Exception as e:
    print(f"💥 Unexpected crash: {e}")
    traceback.print_exc()
finally:
    save_data()
    driver.quit()
    print("🏁 Scraping completed.")