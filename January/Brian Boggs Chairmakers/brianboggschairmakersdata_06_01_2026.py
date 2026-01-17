import time
import pandas as pd
import sys
import signal
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# --- GLOBAL VARIABLES ---
all_scraped_data = []
output_filename = "brianboggs_complete_data.xlsx"

# --- HELPER FUNCTIONS ---
def clean_data(text):
    """Ensures that empty strings or None become 'N/A'."""
    if text is None:
        return "N/A"
    text = str(text).strip()
    if text == "" or text.lower() == "none" or text.lower() == "null":
        return "N/A"
    return text

def save_to_excel(data, filename):
    if not data:
        return
    try:
        df = pd.DataFrame(data)
        # Apply cleaning to the whole dataframe just in case
        df = df.applymap(lambda x: clean_data(x))
        df.to_excel(filename, index=False)
        print(f"\n[SYSTEM] Data saved to '{filename}' ({len(df)} records).")
    except Exception as e:
        print(f"\n[ERROR] Could not save Excel: {e}")

def signal_handler(sig, frame):
    print("\n\n[STOPPING] Script interrupted by user (Ctrl+C). Saving data...")
    save_to_excel(all_scraped_data, output_filename)
    sys.exit(0)

signal.signal(signal.SIGINT, signal_handler)

def get_product_details(driver, product_url, category_context):
    # 1. Initialize ALL fields to "N/A" by default
    details = {
        "Category": "N/A",
        "Product URL": product_url,
        "Product Name": "N/A",
        "SKU": "N/A",
        "Price": "N/A",
        "Brand": "Brian Boggs Chairmakers",
        "Description": "N/A",
        "Height": "N/A",
        "Width": "N/A",
        "Depth": "N/A",
        "Product Sheet": "N/A",
        "Care Instructions": "N/A",
        "Image1": "N/A",
        "Image2": "N/A",
        "Image3": "N/A",
        "Image4": "N/A"
    }

    try:
        driver.get(product_url)
        # Wait for body to ensure page load
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "body")))
        
        # --- 1. CATEGORY ---
        try:
            # User HTML: <nav class="woocommerce-breadcrumb" ...>
            # We try CSS selector first, then xpath as backup
            breadcrumb_el = driver.find_element(By.CSS_SELECTOR, "nav.woocommerce-breadcrumb")
            cat_text = breadcrumb_el.text.strip()
            
            # Fallback: sometimes .text is empty if element is hidden, try get_attribute
            if not cat_text:
                cat_text = breadcrumb_el.get_attribute("innerText")
            
            if cat_text:
                # Format: Home / Tables / ...
                details["Category"] = cat_text.replace("/", " / ").replace("  ", " ").strip()
            else:
                details["Category"] = category_context
        except:
            details["Category"] = category_context

        # --- 2. PRODUCT NAME ---
        try:
            title = driver.find_element(By.CSS_SELECTOR, "h1.product_title").text.strip()
            details["Product Name"] = title
        except:
            pass

        # --- 3. PRICE ---
        try:
            price = driver.find_element(By.CSS_SELECTOR, "p.price").text.strip()
            details["Price"] = price
        except:
            pass

        # --- 4. DESCRIPTION ---
        try:
            desc = driver.find_element(By.CSS_SELECTOR, ".woocommerce-product-details__short-description").text.strip()
            details["Description"] = desc
        except:
            pass

        # --- 5. DIMENSIONS (Height, Width, Depth) ---
        try:
            # Find all widget containers
            widget_containers = driver.find_elements(By.CSS_SELECTOR, ".elementor-widget-container")
            
            target_ps = []
            # Find the specific container having "Height" or "Width"
            for widget in widget_containers:
                if "height" in widget.text.lower() or "width" in widget.text.lower():
                    target_ps = widget.find_elements(By.TAG_NAME, "p")
                    break
            
            if target_ps:
                for i, p in enumerate(target_ps):
                    txt_raw = p.text.strip()
                    txt_lower = txt_raw.lower()

                    # Helper to grab the value (either in same P or next P)
                    def extract_val(current_idx, current_text, label):
                        # Pattern 1: "Height: 45cm"
                        if ":" in current_text:
                            parts = current_text.split(":", 1)
                            if len(parts) > 1 and parts[1].strip():
                                return parts[1].strip()
                        
                        # Pattern 2: "Height" (value is in next P)
                        if current_idx + 1 < len(target_ps):
                            next_text = target_ps[current_idx + 1].text.strip()
                            if next_text:
                                return next_text
                        
                        return "N/A"

                    # Check labels
                    if "height" in txt_lower and "width" not in txt_lower:
                        val = extract_val(i, txt_raw, "Height")
                        if val != "N/A": details["Height"] = val
                    
                    elif "width" in txt_lower:
                        val = extract_val(i, txt_raw, "Width")
                        if val != "N/A": details["Width"] = val
                        
                    elif "depth" in txt_lower:
                        val = extract_val(i, txt_raw, "Depth")
                        if val != "N/A": details["Depth"] = val

        except Exception:
            pass

        # --- 6. FILES ---
        try:
            links = driver.find_elements(By.CSS_SELECTOR, ".elementor-icon-list-item a")
            for l in links:
                h = l.get_attribute("href")
                t = l.text.lower()
                if "product sheet" in t:
                    details["Product Sheet"] = h
                elif "care instructions" in t:
                    details["Care Instructions"] = h
        except:
            pass

        # --- 7. IMAGES ---
        try:
            img_list = []
            # Gather from thumbs first
            thumbs = driver.find_elements(By.CSS_SELECTOR, ".flex-control-nav.flex-control-thumbs li img")
            for t in thumbs:
                src = t.get_attribute("src")
                if src: img_list.append(src)
            
            # If empty, gather from main gallery
            if not img_list:
                mains = driver.find_elements(By.CSS_SELECTOR, ".woocommerce-product-gallery__image")
                for m in mains:
                    try:
                        # try to get the 'href' of the anchor tag inside
                        a = m.find_element(By.TAG_NAME, "a")
                        href = a.get_attribute("href")
                        if href: img_list.append(href)
                    except:
                        pass
            
            # Deduplicate and Assign
            unique = []
            [unique.append(x) for x in img_list if x not in unique]
            
            if len(unique) > 0: details["Image1"] = unique[0]
            if len(unique) > 1: details["Image2"] = unique[1]
            if len(unique) > 2: details["Image3"] = unique[2]
            if len(unique) > 3: details["Image4"] = unique[3]

        except:
            pass

    except Exception as e:
        print(f" [Error on {product_url}]: {e}")

    # FINAL CLEANUP: Ensure every single value is "N/A" if empty
    for k, v in details.items():
        details[k] = clean_data(v)

    return details

def scrape_brianboggs_final():
    global all_scraped_data

    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    categories = {
        "Seating": "https://brianboggschairmakers.com/seating/",
        "Tables": "https://brianboggschairmakers.com/tables/",
        "Furniture for Musicians": "https://brianboggschairmakers.com/furniture-for-musicians/",
        "Desks & Office": "https://brianboggschairmakers.com/desks-office/",
        "Display & Storage": "https://brianboggschairmakers.com/display-storage/",
        "Current Inventory": "https://brianboggschairmakers.com/shop-products/"
    }
    
    # 1. Collect URLs
    product_queue = []
    seen_urls = set()

    print("--- PHASE 1: Collecting URLs ---")
    for cat_name, cat_url in categories.items():
        try:
            driver.get(cat_url)
            time.sleep(2)
            
            # Load More Loop
            while True:
                try:
                    btn = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Load More')]")))
                    driver.execute_script("arguments[0].click();", btn)
                    time.sleep(2)
                except:
                    break
            
            cards = driver.find_elements(By.CSS_SELECTOR, "li.product")
            for c in cards:
                try:
                    # Finds the link
                    lnk = c.find_element(By.TAG_NAME, "a").get_attribute("href")
                    # Finds the title
                    try:
                        nm = c.find_element(By.CSS_SELECTOR, ".woocommerce-loop-product__title").text.strip()
                    except:
                        nm = "Unknown"
                        
                    if lnk and lnk not in seen_urls:
                        seen_urls.add(lnk)
                        product_queue.append({"url": lnk, "name": nm, "cat": cat_name})
                except:
                    continue
            print(f" > {cat_name}: Found {len(product_queue)} total so far.")
            
        except Exception as e:
            print(f"Error in {cat_name}: {e}")

    # 2. Extract Details
    total = len(product_queue)
    print(f"\n--- PHASE 2: Scraping {total} Products ---")
    
    for idx, item in enumerate(product_queue):
        print(f"Scraping {idx+1}/{total} -> {item['name']} -> {item['url']}")
        
        data = get_product_details(driver, item['url'], item['cat'])
        all_scraped_data.append(data)
        
        if len(all_scraped_data) % 50 == 0:
            print("[Saving Batch...]")
            save_to_excel(all_scraped_data, output_filename)
            
    driver.quit()
    save_to_excel(all_scraped_data, output_filename)
    print("\nDone.")

if __name__ == "__main__":
    scrape_brianboggs_final()