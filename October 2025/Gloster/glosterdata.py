# import time
# import pandas as pd
# import signal
# import sys
# import traceback
# from selenium import webdriver
# from selenium.webdriver.common.by import By
# from selenium.webdriver.chrome.service import Service
# from selenium.webdriver.chrome.options import Options
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
# from webdriver_manager.chrome import ChromeDriverManager

# # ==================================== CONFIG ====================================
# BASE_URL = "https://www.gloster.com/en"
# OUTPUT_FILE = "gloster_products.xlsx"
# SAVE_INTERVAL = 50  # Save data every 50 products

# # ==================================== SETUP SELENIUM ====================================
# options = Options()
# options.add_argument("--start-maximized")
# options.add_argument("--disable-blink-features=AutomationControlled")
# driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
# wait = WebDriverWait(driver, 20)

# # ==================================== DATA STORAGE ====================================
# data = []

# def save_data():
#     """Save collected data to Excel safely."""
#     if not data:
#         return
#     df = pd.DataFrame(data)
#     df.drop_duplicates(inplace=True)
#     df.to_excel(OUTPUT_FILE, index=False)
#     print(f"💾 Progress saved! ({len(df)} total records)")

# # Handle Ctrl+C or crash
# def handle_exit(sig, frame):
#     global stop_requested
#     print("\n🛑 Ctrl+C detected! Saving progress before exit...")
#     stop_requested = True
#     save_data()
#     driver.quit()
#     raise SystemExit("✅ Exited safely after saving progress.")

# signal.signal(signal.SIGINT, handle_exit)

# # ==================================== UTILS ====================================
# def safe_get_text(el):
#     try:
#         return el.text.strip()
#     except:
#         return ""

# def extract_breadcrumbs():
#     crumbs = []
#     try:
#         li_items = driver.find_elements(By.CSS_SELECTOR, "breadcrumbs ul.breadcrumbs li.breadcrumbs__item")
#         for li in li_items:
#             label = li.find_element(By.XPATH, ".//label|.//a")
#             crumbs.append(label.text.strip())
#     except:
#         pass
#     return " > ".join(crumbs)

# # def extract_attributes():
# #     """Extract product dimensions by expanding the 'Dimensions' dropdown."""
# #     attrs = {
# #         "Attr_Width": "", "Attr_SeatHeight": "", "Attr_Height": "", "Attr_Depth": "",
# #         "Attr_ArmHeight": "", "Attr_Length": "", "Attr_CubicSize": "", "Attr_Weight": ""
# #     }
# #     try:
# #         # Try clicking the 'Dimensions' dropdown
# #         dim_button = driver.find_element(By.XPATH, "//h2[normalize-space()='Dimensions']/following-sibling::button")
# #         driver.execute_script("arguments[0].scrollIntoView(true);", dim_button)
# #         time.sleep(1)
# #         dim_button.click()
# #         time.sleep(1)

# #         # Now extract the visible attribute list
# #         items = driver.find_elements(By.XPATH, "//h2[normalize-space()='Dimensions']/following::div[contains(@class,'attribute-list__items')]//app-attribute-list-item")
# #         for item in items:
# #             title = safe_get_text(item.find_element(By.CSS_SELECTOR, ".attribute-list-item__title span"))
# #             value = safe_get_text(item.find_element(By.CSS_SELECTOR, ".attribute-list-item__value span"))
# #             # ✅ Important: match more specific attributes first
# #             if "Arm Height" in title:
# #                 attrs["Attr_ArmHeight"] = value
# #             elif "Seat Height" in title:
# #                 attrs["Attr_SeatHeight"] = value
# #             elif "Width" in title:
# #                 attrs["Attr_Width"] = value
# #             elif "Height" in title:
# #                 attrs["Attr_Height"] = value
# #             elif "Depth" in title:
# #                 attrs["Attr_Depth"] = value
# #             elif "Length" in title:
# #                 attrs["Attr_Length"] = value
# #             elif "Cubic Size" in title:
# #                 attrs["Attr_CubicSize"] = value
# #             elif "Weight" in title:
# #                 attrs["Attr_Weight"] = value
# #     except Exception as e:
# #         print("⚠️ Dimensions not found:", e)
# #     return attrs

# # def extract_spec_sheets():
# #     """Extract PDF links by expanding the 'Downloads' dropdown."""
# #     sheets = {
# #         "Spec Sheet": "", "Assembly Instructions": "", "Warranty": "",
# #         "Outdoor Fabrics Care Sheet": "", "Outdoor Rope Care Sheet": "",
# #         "Powder Coated Aluminium Care Sheet": "", "Brushed Stainless Steel Care Sheet": "",
# #         "Sling Care Sheet": "", "Teak Care Sheet": "", "Wicker Care Sheet": "",
# #         "Protective Covers Care Sheet": ""
# #     }
# #     try:
# #         # Try clicking the 'Downloads' dropdown
# #         dl_button = driver.find_element(By.XPATH, "//h2[normalize-space()='Downloads']/following-sibling::button")
# #         driver.execute_script("arguments[0].scrollIntoView(true);", dl_button)
# #         time.sleep(1)
# #         dl_button.click()
# #         time.sleep(1)

# #         # Extract the visible PDF items
# #         items = driver.find_elements(By.XPATH, "//h2[normalize-space()='Downloads']/following::div[contains(@class,'attribute-list__items')]//app-attribute-list-item")
# #         for item in items:
# #             title = safe_get_text(item.find_element(By.CSS_SELECTOR, ".attribute-list-item__title span"))
# #             try:
# #                 link = item.find_element(By.CSS_SELECTOR, ".attribute-list-item__value a").get_attribute("href")
# #             except:
# #                 link = ""
# #             if title in sheets:
# #                 sheets[title] = link
# #     except Exception as e:
# #         print("⚠️ Downloads not found:", e)
# #     return sheets

# def extract_attributes():
#     """Extract product dimensions by expanding the 'Dimensions' dropdown."""
#     attrs = {
#         "Attr_Width": "", "Attr_SeatHeight": "", "Attr_Height": "", "Attr_Depth": "",
#         "Attr_ArmHeight": "", "Attr_Length": "", "Attr_CubicSize": "", "Attr_Weight": "","Attr_Diameter": "",
#         "Attr_Clearance_Under_Table": ""
#     }

#     # Try clicking the 'Dimensions' dropdown if it exists
#     dim_buttons = driver.find_elements(By.XPATH, "//h2[normalize-space()='Dimensions']/following-sibling::button")
#     if dim_buttons:
#         dim_button = dim_buttons[0]
#         driver.execute_script("arguments[0].scrollIntoView(true);", dim_button)
#         time.sleep(1)
#         dim_button.click()
#         time.sleep(1)

#         # Extract attribute items
#         items = driver.find_elements(By.XPATH, "//h2[normalize-space()='Dimensions']/following::div[contains(@class,'attribute-list__items')]//app-attribute-list-item")
#         for item in items:
#             title_elems = item.find_elements(By.CSS_SELECTOR, ".attribute-list-item__title span")
#             value_elems = item.find_elements(By.CSS_SELECTOR, ".attribute-list-item__value span")
#             if not title_elems or not value_elems:
#                 continue  # skip if either missing
#             title = safe_get_text(title_elems[0])
#             value = safe_get_text(value_elems[0])

#             # Match more specific first
#             if "Arm Height" in title:
#                 attrs["Attr_ArmHeight"] = value
#             elif "Seat Height" in title:
#                 attrs["Attr_SeatHeight"] = value
#             elif "Width" in title:
#                 attrs["Attr_Width"] = value
#             elif "Height" in title:
#                 attrs["Attr_Height"] = value
#             elif "Depth" in title:
#                 attrs["Attr_Depth"] = value
#             elif "Length" in title:
#                 attrs["Attr_Length"] = value
#             elif "Cubic Size" in title:
#                 attrs["Attr_CubicSize"] = value
#             elif "Weight" in title:
#                 attrs["Attr_Weight"] = value
#             elif "Diameter" in title:
#                 attrs["Attr_Diameter"] = value
#             elif "Clearance Under Table" in title:
#                 attrs["Attr_Clearance_Under_Table"] = value
#     # No print if dropdown missing, just return empty attributes
#     return attrs


# def extract_spec_sheets():
#     """Extract PDF links by expanding the 'Downloads' dropdown."""
#     sheets = {
#         "Spec Sheet": "", "Assembly Instructions": "", "Warranty": "",
#         "Outdoor Fabrics Care Sheet": "", "Outdoor Rope Care Sheet": "",
#         "Powder Coated Aluminium Care Sheet": "", "Brushed Stainless Steel Care Sheet": "",
#         "Sling Care Sheet": "", "Teak Care Sheet": "", "Wicker Care Sheet": "",
#         "Protective Covers Care Sheet": ""
#     }

#     dl_buttons = driver.find_elements(By.XPATH, "//h2[normalize-space()='Downloads']/following-sibling::button")
#     if dl_buttons:
#         dl_button = dl_buttons[0]
#         driver.execute_script("arguments[0].scrollIntoView(true);", dl_button)
#         time.sleep(1)
#         dl_button.click()
#         time.sleep(1)

#         items = driver.find_elements(By.XPATH, "//h2[normalize-space()='Downloads']/following::div[contains(@class,'attribute-list__items')]//app-attribute-list-item")
#         for item in items:
#             title_elems = item.find_elements(By.CSS_SELECTOR, ".attribute-list-item__title span")
#             value_links = item.find_elements(By.CSS_SELECTOR, ".attribute-list-item__value a")
#             if not title_elems:
#                 continue
#             title = safe_get_text(title_elems[0])
#             link = value_links[0].get_attribute("href") if value_links else ""
#             if title in sheets:
#                 sheets[title] = link
#     # No print if dropdown missing
#     return sheets

# def extract_images():
#     images = []
#     try:
#         imgs = driver.find_elements(By.CSS_SELECTOR, "div.media__spinner-container img.media__spinner-image")
#         for img in imgs[:4]:  # get max 4 images
#             images.append(img.get_attribute("src"))
#     except:
#         pass
#     # Pad with empty strings if less than 4 images
#     while len(images) < 4:
#         images.append("")
#     return images[:4]

# # ==================================== SCRAPER ====================================
# try:
#     print("🔍 Opening main site...")
#     driver.get(BASE_URL)
#     time.sleep(5)

#     # Open menu -> products -> collections
#     print("📂 Opening MENU...")
#     wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.navigation-trigger"))).click()
#     time.sleep(2)
#     print("📦 Clicking on 'Products'...")
#     wait.until(EC.element_to_be_clickable((By.XPATH, "//a[normalize-space()='Products']"))).click()
#     time.sleep(2)
#     print("🗂️ Clicking on 'Collections'...")
#     wait.until(EC.element_to_be_clickable((By.XPATH, "//a[normalize-space()='Collections']"))).click()
#     time.sleep(5)

#     # Extract collection links
#     collection_links = list({el.get_attribute("href") for el in driver.find_elements(By.XPATH, "//a[contains(@href, '/en/products/collections/')]")})
#     print(f"✅ Found {len(collection_links)} collections.")

#     # Visit collections
#     for idx, cat_url in enumerate(collection_links, start=1):
#         print(f"\n==============================")
#         print(f"[{idx}/{collection_links}] 🏷️ Visiting Collection: {cat_url}")
#         print(f"==============================")

#         driver.get(cat_url)
#         time.sleep(5)
#         driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
#         time.sleep(3)

#         product_links = list({p.get_attribute("href") for p in driver.find_elements(By.XPATH, "//a[contains(@href, '/en/products/collections/')]") if p.get_attribute("href").count("/") > 6})

#         print(f"📦 Found {product_links} products in this collection.")

#         for i, product_url in enumerate(product_links, start=1):
#             driver.get(product_url)
#             time.sleep(5)

#             breadcrumbs = extract_breadcrumbs()
#             try:
#                 product_name = driver.find_element(By.CSS_SELECTOR, "section.product__header h1").text.strip() + " " + driver.find_element(By.CSS_SELECTOR, "section.product__header h2").text.strip()
#             except:
#                 product_name = ""

#             attrs = extract_attributes()
#             specs = extract_spec_sheets()
#             images = extract_images()
#             try:
#                 description_html = driver.find_element(By.CSS_SELECTOR, "section product-attributes-group").get_attribute("innerHTML")
#             except:
#                 description_html = ""

#             product_data = {
#                 "Category Breadcrumbs": breadcrumbs,
#                 "Product URL": product_url,
#                 "Product Name": product_name,
#                 "SKU": "",
#                 "Brand": "GLOSTER",
#                 "Full Description HTML": description_html,
#                 **attrs,
#                 **specs,
#                 "Image1": images[0], "Image2": images[1], "Image3": images[2], "Image4": images[3]
#             }

#             data.append(product_data)
#             print(f"   🔸 [{i}/{len(product_links)}] {product_url} -> {product_name}")

#             # Autosave
#             if len(data) % SAVE_INTERVAL == 0:
#                 save_data()

# except Exception as e:
#     print(f"\n❌ Error occurred: {e}")
#     save_data()

# finally:
#     # Final save
#     save_data()
#     driver.quit()
#     print("👋 Browser closed.")



import time
import pandas as pd
import signal
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ==================================== CONFIG ====================================
BASE_URL = "https://www.gloster.com/en"
OUTPUT_FILE = "gloster_products.xlsx"
SAVE_INTERVAL = 50  # Save every 50 products

# ==================================== SELENIUM SETUP ====================================
options = Options()
options.add_argument("--start-maximized")
options.add_argument("--disable-blink-features=AutomationControlled")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 20)

# ==================================== DATA STORAGE ====================================
data = []
stop_requested = False

def save_data():
    """Save collected data to Excel safely."""
    if not data:
        return
    df = pd.DataFrame(data)
    df.drop_duplicates(inplace=True)
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"💾 Progress saved! ({len(df)} total records)")

# Handle Ctrl+C
def handle_exit(sig, frame):
    global stop_requested
    print("\n🛑 Ctrl+C detected! Saving progress before exit...")
    stop_requested = True
    save_data()
    driver.quit()
    sys.exit("✅ Exited safely after saving progress.")

signal.signal(signal.SIGINT, handle_exit)

# ==================================== UTILS ====================================
def safe_get_text(el):
    try:
        return el.text.strip()
    except:
        return ""

def extract_breadcrumbs():
    crumbs = []
    try:
        li_items = driver.find_elements(By.CSS_SELECTOR, "breadcrumbs ul.breadcrumbs li.breadcrumbs__item")
        for li in li_items:
            label = li.find_element(By.XPATH, ".//label|.//a")
            crumbs.append(label.text.strip())
    except:
        pass
    return " > ".join(crumbs)

def extract_attributes():
    attrs = {
        "Attr_Width": "", "Attr_SeatHeight": "", "Attr_Height": "", "Attr_Depth": "",
        "Attr_ArmHeight": "", "Attr_Length": "", "Attr_CubicSize": "", "Attr_Weight": "",
        "Attr_Diameter": "", "Attr_Clearance_Under_Table": ""
    }
    dim_buttons = driver.find_elements(By.XPATH, "//h2[normalize-space()='Dimensions']/following-sibling::button")
    if dim_buttons:
        dim_button = dim_buttons[0]
        driver.execute_script("arguments[0].scrollIntoView(true);", dim_button)
        time.sleep(1)
        dim_button.click()
        time.sleep(1)
        items = driver.find_elements(By.XPATH, "//h2[normalize-space()='Dimensions']/following::div[contains(@class,'attribute-list__items')]//app-attribute-list-item")
        for item in items:
            title_elems = item.find_elements(By.CSS_SELECTOR, ".attribute-list-item__title span")
            value_elems = item.find_elements(By.CSS_SELECTOR, ".attribute-list-item__value span")
            if not title_elems or not value_elems:
                continue
            title = safe_get_text(title_elems[0])
            value = safe_get_text(value_elems[0])
            if "Arm Height" in title:
                attrs["Attr_ArmHeight"] = value
            elif "Seat Height" in title:
                attrs["Attr_SeatHeight"] = value
            elif "Width" in title:
                attrs["Attr_Width"] = value
            elif "Height" in title:
                attrs["Attr_Height"] = value
            elif "Depth" in title:
                attrs["Attr_Depth"] = value
            elif "Length" in title:
                attrs["Attr_Length"] = value
            elif "Cubic Size" in title:
                attrs["Attr_CubicSize"] = value
            elif "Weight" in title:
                attrs["Attr_Weight"] = value
            elif "Diameter" in title:
                attrs["Attr_Diameter"] = value
            elif "Clearance Under Table" in title:
                attrs["Attr_Clearance_Under_Table"] = value
    return attrs

def extract_spec_sheets():
    sheets = {
        "Spec Sheet": "", "Assembly Instructions": "", "Warranty": "",
        "Outdoor Fabrics Care Sheet": "", "Outdoor Rope Care Sheet": "",
        "Powder Coated Aluminium Care Sheet": "", "Brushed Stainless Steel Care Sheet": "",
        "Sling Care Sheet": "", "Teak Care Sheet": "", "Wicker Care Sheet": "",
        "Protective Covers Care Sheet": ""
    }
    dl_buttons = driver.find_elements(By.XPATH, "//h2[normalize-space()='Downloads']/following-sibling::button")
    if dl_buttons:
        dl_button = dl_buttons[0]
        driver.execute_script("arguments[0].scrollIntoView(true);", dl_button)
        time.sleep(1)
        dl_button.click()
        time.sleep(1)
        items = driver.find_elements(By.XPATH, "//h2[normalize-space()='Downloads']/following::div[contains(@class,'attribute-list__items')]//app-attribute-list-item")
        for item in items:
            title_elems = item.find_elements(By.CSS_SELECTOR, ".attribute-list-item__title span")
            value_links = item.find_elements(By.CSS_SELECTOR, ".attribute-list-item__value a")
            if not title_elems:
                continue
            title = safe_get_text(title_elems[0])
            link = value_links[0].get_attribute("href") if value_links else ""
            if title in sheets:
                sheets[title] = link
    return sheets

def extract_images():
    images = []
    try:
        imgs = driver.find_elements(By.CSS_SELECTOR, "div.media__spinner-container img.media__spinner-image")
        for img in imgs[:4]:
            images.append(img.get_attribute("src"))
    except:
        pass
    while len(images) < 4:
        images.append("")
    return images[:4]

# ==================================== SCRAPER ====================================
try:
    print("🔍 Opening main site...")
    driver.get(BASE_URL)
    time.sleep(5)

    # Navigate to Collections
    print("📂 Opening MENU...")
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.navigation-trigger"))).click()
    time.sleep(2)
    print("📦 Clicking on 'Products'...")
    wait.until(EC.element_to_be_clickable((By.XPATH, "//a[normalize-space()='Products']"))).click()
    time.sleep(2)
    print("🗂️ Clicking on 'Collections'...")
    wait.until(EC.element_to_be_clickable((By.XPATH, "//a[normalize-space()='Collections']"))).click()
    time.sleep(5)

    # Extract collection links
    collection_links = list({el.get_attribute("href") for el in driver.find_elements(By.XPATH, "//a[contains(@href, '/en/products/collections/')]")})
    print(f"✅ Found {len(collection_links)} collections.")

    # Visit collections
    for idx, cat_url in enumerate(collection_links, start=1):
        if stop_requested: break
        print(f"\n==============================")
        print(f"[{idx}/{len(collection_links)}] 🏷️ Visiting Collection: {cat_url}")
        print(f"==============================")

        driver.get(cat_url)
        time.sleep(5)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(3)

        # Extract product links
        product_links = list({p.get_attribute("href") for p in driver.find_elements(By.XPATH, "//a[contains(@href, '/en/products/collections/')]") if p.get_attribute("href").count("/") > 6})
        print(f"📦 Found {len(product_links)} products in this collection.")

        for i, product_url in enumerate(product_links, start=1):
            if stop_requested: break
            driver.get(product_url)
            time.sleep(5)

            breadcrumbs = extract_breadcrumbs()
            try:
                product_name = driver.find_element(By.CSS_SELECTOR, "section.product__header h1").text.strip() + " " + driver.find_element(By.CSS_SELECTOR, "section.product__header h2").text.strip()
            except:
                product_name = ""

            attrs = extract_attributes()
            specs = extract_spec_sheets()
            images = extract_images()
            try:
                description_html = driver.find_element(By.CSS_SELECTOR, "section product-attributes-group").get_attribute("innerHTML")
            except:
                description_html = ""

            product_data = {
                "Category Breadcrumbs": breadcrumbs,
                "Product URL": product_url,
                "Product Name": product_name,
                "SKU": "",
                "Brand": "GLOSTER",
                "Full Description HTML": description_html,
                **attrs,
                **specs,
                "Image1": images[0], "Image2": images[1], "Image3": images[2], "Image4": images[3]
            }

            data.append(product_data)
            print(f"   🔸 [{i}/{len(product_links)}] {product_name} -> {product_url}")

            if len(data) % SAVE_INTERVAL == 0:
                save_data()

except Exception as e:
    print(f"\n❌ Error occurred: {e}")
    # save_data()

finally:
    save_data()
    driver.quit()
    print("👋 Browser closed.")
