# coding: utf-8
import time
import re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager

# ---------- CONFIG ----------
OUTPUT_FILE = "domiziani_products_details.xlsx"
BRAND = "Domiziani-America"

PRODUCT_URLS = {
    "Table Tops": [
        "https://www.domizianiamerica.com/product-page/yellowstone",
        "https://www.domizianiamerica.com/product-page/arizona",
        "https://www.domizianiamerica.com/product-page/lanzarote",
        "https://www.domizianiamerica.com/product-page/forma-nero",
        "https://www.domizianiamerica.com/product-page/terra-ikebana",
        "https://www.domizianiamerica.com/product-page/terra-forma-zaffiro",
        "https://www.domizianiamerica.com/product-page/aria-nembo",
        "https://www.domizianiamerica.com/product-page/acqua-bajkal",
        "https://www.domizianiamerica.com/product-page/terra-roccia",
        "https://www.domizianiamerica.com/product-page/aria-tramonto",
        "https://www.domizianiamerica.com/product-page/roccia-turchese-1",
        "https://www.domizianiamerica.com/product-page/terra-masseo",
        "https://www.domizianiamerica.com/product-page/terra-roccia",
        "https://www.domizianiamerica.com/product-page/terra-positano-notte",
        "https://www.domizianiamerica.com/product-page/blu-cava",
        "https://www.domizianiamerica.com/product-page/luna-rosa",
        "https://www.domizianiamerica.com/product-page/pavone-blu",
        "https://www.domizianiamerica.com/product-page/sun-141-red"
    ],
    "Table Sets": [
        "https://www.domizianiamerica.com/product-page/octopus-table-set",
        "https://www.domizianiamerica.com/product-page/brando-table-set"
    ],
    "Fire Pits": [
        "https://www.domizianiamerica.com/product-page/fire-pits"
    ]
}

# ---------- SETUP CHROME ----------
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_argument("--disable-notifications")

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

def extract_full_accordion(driver):
    """Extracts the full accordion HTML block (Pattern Description, Available Sizes, etc.)."""
    try:
        accordion = driver.find_element(
            By.CSS_SELECTOR,
            "div[id*='comp-lkps559k_r_comp-l2itfe7i_r_comp-kquuuy79']"
        )
        return accordion.get_attribute("outerHTML")
    except NoSuchElementException:
        return ""

# ---------- SCRAPE PRODUCT DETAILS ----------
def scrape_product(cat_name, product_url):
    data = {
        "Category": cat_name,
        "Product_URL": product_url,
        "Product_Name": "",
        "SKU": "",
        "Brand": BRAND,
        "Description": "",
        "Full_Description_HTML": "",
        "Image1": "",
        "Image2": "",
        "Image3": "",
        "Image4": "",
    }

    try:
        driver.get(product_url)
        time.sleep(3)  # wait for page load

        # Product name
        try:
            name_el = driver.find_element(By.CSS_SELECTOR, "h1.font_0.wixui-rich-text__text")
            data["Product_Name"] = name_el.text.strip()
        except NoSuchElementException:
            pass

        # Product images (up to 4)
        try:
            img_els = driver.find_elements(By.CSS_SELECTOR, "div[data-hook='image-item'] img")
            img_urls = []
            for img in img_els:
                src = img.get_attribute("src")
                if src and src not in img_urls:
                    img_urls.append(src)
            for i in range(4):
                if i < len(img_urls):
                    data[f"Image{i+1}"] = img_urls[i]
        except NoSuchElementException:
            pass

        # ---------- SCRAPE SKU & DESCRIPTION ----------
        try:
            # Select the rich-text div containing description + SKU
            rich_div = driver.find_element(
                By.CSS_SELECTOR,
                "div[class*='comp-lkps559k_r_comp-l2itfe7i_r_comp-kwezj0j5']"
            )

            # Save full HTML
            data["Full_Description_HTML"] = extract_full_accordion(driver)

            # Extract all <p> tags
            p_tags = rich_div.find_elements(By.TAG_NAME, "p")
            description_texts = []

            for p in p_tags:
                text = p.text.strip()
                if text:
                    description_texts.append(text)

            # For Table Tops:
            if cat_name == "Table Tops":
                if len(description_texts) >= 2:
                    # All except last <p> as description (joined)
                    data["Description"] = " ".join(description_texts[:-1])
                    # Last <p> as SKU
                    data["SKU"] = description_texts[-1]
                elif len(description_texts) == 1:
                    data["Description"] = description_texts[0]

            # For Table Sets / Fire Pits:
            else:
                if description_texts:
                    data["Description"] = description_texts[0]
                # SKU might not exist in these, so keep blank if not found

        except NoSuchElementException:
            pass


        print(f"✅ {cat_name} | {data['Product_Name']} ({data['SKU']})")
        return data

    except Exception as e:
        print(f"⚠️ Error scraping {product_url} | {e}")
        return data


# ---------- MAIN ----------
all_data = []

for cat_name, urls in PRODUCT_URLS.items():
    for url in urls:
        product_data = scrape_product(cat_name, url)
        all_data.append(product_data)

# Save to Excel
if all_data:
    df = pd.DataFrame(all_data)
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"\n✅ Total {len(all_data)} products saved to {OUTPUT_FILE}")
else:
    print("\n⚠️ No products scraped.")

driver.quit()
