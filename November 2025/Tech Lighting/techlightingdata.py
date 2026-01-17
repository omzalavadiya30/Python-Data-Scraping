import pandas as pd
import time
import signal
import sys
from urllib.parse import urljoin, unquote
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ---------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------
OUTPUT_FILE = "techlighting_final_complete.xlsx"
SAVE_BATCH_SIZE = 50

# Full list of Categories to scan
CATEGORY_URLS = [
    "https://www.techlighting.com/Designers/Collections/Avroko",
    "https://www.techlighting.com/Designers/Collections/Clodagh",
    "https://www.techlighting.com/Designers/Collections/Kelly-Wearstler",
    "https://www.techlighting.com/Designers/Collections/Mick-De-Giulio",
    "https://www.techlighting.com/Designers/Collections/Sean-Lavin",
    "https://www.techlighting.com/Products/AllFixtures?filter=new",
    "https://www.techlighting.com/Products/AllFixtures",
    "https://www.techlighting.com/Products/Fixtures/Chandeliers",
    "https://www.techlighting.com/Products/Fixtures/Pendants",
    "https://www.techlighting.com/Products/Fixtures/Linear-Suspension",
    "https://www.techlighting.com/Products/Fixtures/Table",
    "https://www.techlighting.com/Products/Fixtures/Floor",
    "https://www.techlighting.com/Products/Fixtures/Flush-Mounts",
    "https://www.techlighting.com/Products/Fixtures/Bath-Collection",
    "https://www.techlighting.com/Products/Fixtures/Wall-Collection",
    "https://www.techlighting.com/Products/Fixtures/Low-Voltage-Heads",
    "https://www.techlighting.com/Products/Fixtures/Multi-Port-Chandeliers",
    "https://www.techlighting.com/Products/Fixtures/Unilume-LED-Accent-And-Task",
    "https://www.techlighting.com/Products/Fixtures/Accessories",
    "https://www.techlighting.com/Products/AllFixtures?locationcategory=systems",
    "https://www.techlighting.com/Products/Brox",
    "https://www.techlighting.com/Products/Essence",
    "https://www.techlighting.com/Products/Hardware/Monorail",
    "https://www.techlighting.com/Products/Hardware/Kable-Lite",
    "https://www.techlighting.com/Products/Hardware/Wall-Monorail",
    "https://www.techlighting.com/Products/Hardware/Freejack",
    "https://www.techlighting.com/Products/AllFixtures?filter=new&locationcategory=outdoor",
    "https://www.techlighting.com/Products/AllFixtures?locationcategory=outdoor",
    "https://www.techlighting.com/Products/Fixtures/Outdoor-Bollards",
    "https://www.techlighting.com/Products/Fixtures/Outdoor-Flush-Mounts",
    "https://www.techlighting.com/Products/Fixtures/Outdoor-Light-Columns",
    "https://www.techlighting.com/Products/Fixtures/Outdoor-Pathway-Lights",
    "https://www.techlighting.com/Products/Fixtures/Outdoor-Step-Lights",
    "https://www.techlighting.com/Products/Fixtures/Outdoor-Pendant-Lights",
    "https://www.techlighting.com/Products/Fixtures/Outdoor-Wall-Sconces",
    "https://www.techlighting.com/Products/Fixtures/Outdoor-Portables",
    "https://v1.element-lighting.com/Products/#Element-Downlights",
    "https://v1.element-lighting.com/Products#Entra-Downlights",
    "https://v1.element-lighting.com/Products#Element-Cylinders",
    "https://v1.element-lighting.com/Products#Entra-Cylinders",
    "https://v1.element-lighting.com/Products#VERSE-3-LED-Fixed-Downlight-and-Wall-Wash",
    "https://v1.element-lighting.com/Products#Reflections-Decorative-Downlights",
    "https://v1.element-lighting.com/Products#Multiples"
]

extracted_data = []
product_data_cache = {} 

def setup_driver():
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--log-level=3")
    options.page_load_strategy = 'eager' 
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver

def save_data(data, filename):
    if not data: return
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)
    print(f"\n[SYSTEM] Saved {len(data)} rows to {filename}")

def fix_url(base_url, relative_url):
    if not relative_url or relative_url == "N/A": return "N/A"
    return urljoin(base_url, relative_url)

def scroll_to_element(driver, element):
    try:
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element)
        time.sleep(0.5) 
    except: pass

def handle_accordion(driver, header_id, content_id):
    """
    Finds header, scrolls to it, clicks it if content is not shown.
    """
    try:
        header = driver.find_element(By.ID, header_id)
        content = driver.find_element(By.ID, content_id)
        
        # Scroll to header so click is registered
        scroll_to_element(driver, header)
        
        # If class doesn't contain 'show', click to open
        if "show" not in content.get_attribute("class"):
            driver.execute_script("arguments[0].click();", header)
            time.sleep(1) # Wait for animation
    except: pass

def get_safe_text(driver, css_selector):
    try:
        elem = driver.find_element(By.CSS_SELECTOR, css_selector)
        return driver.execute_script("return arguments[0].textContent;", elem).strip()
    except: return "N/A"

# ---------------------------------------------------------
# PHASE 1: COLLECT ALL URLs (ALLOWING DUPLICATES)
# ---------------------------------------------------------
def collect_products(driver):
    print("--- PHASE 1: Collecting Product URLs (Including Duplicates) ---")
    product_tasks = [] 
    total = len(CATEGORY_URLS)
    
    for idx, url in enumerate(CATEGORY_URLS):
        try:
            raw_name = url.split("?")[0].split("#")[0].split("/")[-1].replace("-", " ")
            if not raw_name: raw_name = "General"
            category_name = unquote(raw_name)
        except:
            category_name = "Unknown"

        print(f"[{idx+1}/{total}] Scanning: {category_name}")
        
        max_retries = 3
        for attempt in range(max_retries):
            try:
                driver.get(url)
                time.sleep(3)
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(2)
                
                grid_links = driver.find_elements(By.CSS_SELECTOR, "a.results-link")
                list_links = driver.find_elements(By.CSS_SELECTOR, "div.indoorSubSection ul li a")
                all_potential = grid_links + list_links
                
                count = 0
                page_unique = set()

                for link in all_potential:
                    try:
                        href = link.get_attribute("href")
                        if not href: continue
                        href = href.split("?")[0].split("#")[0]
                        full_url = fix_url(driver.current_url, href)

                        if "/Products/" not in full_url and "/Fixtures/" not in full_url: continue
                        if any(x in full_url for x in ["sort=", "filter=", "view=", "direction="]): continue
                        if href.endswith("#"): continue
                        if full_url == url: continue

                        if full_url not in page_unique:
                            page_unique.add(full_url)
                            product_tasks.append({
                                "Category Name": category_name,
                                "Product URL": full_url
                            })
                            count += 1
                    except: continue
                
                print(f"   -> Added {count} products to processing list.")
                break 

            except Exception as e:
                print(f"   -> Error on attempt {attempt+1}: {e}")
                if attempt < max_retries - 1:
                    time.sleep(5)
                    driver.refresh()
                else:
                    print("   -> Skipping category due to repeated failures.")

    return product_tasks

# ---------------------------------------------------------
# PHASE 2: SMART SCRAPE (WITH CACHE)
# ---------------------------------------------------------
def scrape_details(driver, tasks):
    print(f"\n--- PHASE 2: Processing {len(tasks)} Items (Using Smart Cache) ---")
    global extracted_data, product_data_cache
    
    for i, task in enumerate(tasks):
        p_url = task['Product URL']
        cat_name = task['Category Name']
        
        if p_url in product_data_cache:
            print(f"[{i+1}/{len(tasks)}] (Cached) Mapping to: {cat_name}")
            cached_row = product_data_cache[p_url].copy()
            cached_row["Category Hierarchy"] = cat_name 
            extracted_data.append(cached_row)
            continue
        
        try:
            driver.get(p_url)
            
            try:
                WebDriverWait(driver, 4).until(EC.presence_of_element_located((By.TAG_NAME, "h1")))
                p_name = driver.find_element(By.TAG_NAME, "h1").text.strip()
            except:
                print(f"   [SKIP] Timeout: {p_url}")
                continue

            if p_name.upper() in ["INDOOR", "ESSENCE", "ELEMENT", "ENTRA", "OUTDOOR"]: continue

            print(f"[{i+1}/{len(tasks)}] (Scraping) -> {p_name}")

            # SKU
            sku = "N/A"
            try:
                dropdowns = driver.find_elements(By.CLASS_NAME, "inb_SkuDropDown")
                for dd in dropdowns:
                    if "MODEL" in dd.text.upper() and "TYPE" not in dd.text.upper():
                        sku = dd.text.replace("MODEL", "").strip()
                        break
                if sku == "N/A":
                    sku = driver.find_element(By.CLASS_NAME, "product-basecode").text.strip()
            except: pass

            # ---------------------------------------------------------
            # HANDLE ACCORDIONS
            # ---------------------------------------------------------
            handle_accordion(driver, "pdp-header-description", "collapse-description")
            handle_accordion(driver, "pdp-header-lamping", "collapse-lamping")
            handle_accordion(driver, "pdp-header-dimensions", "collapse-dimensions")
            # NEW: Handle Downloads accordion
            handle_accordion(driver, "pdp-header-downloads", "collapse-downloads")

            # Extract Text Content
            desc = get_safe_text(driver, "#collapse-description .card-body")
            lamping = get_safe_text(driver, "#collapse-lamping .card-body")

            # ---------------------------------------------------------
            # SPEC SHEET EXTRACTION (From Downloads Accordion)
            # ---------------------------------------------------------
            spec_sheet = "N/A"
            try:
                # Look inside the #collapse-downloads container for a link with text "Spec Sheet"
                downloads_container = driver.find_element(By.ID, "collapse-downloads")
                # Find links inside this specific container
                links = downloads_container.find_elements(By.TAG_NAME, "a")
                
                for link in links:
                    if "Spec Sheet" in link.text:
                        raw_href = link.get_attribute("href")
                        spec_sheet = fix_url(driver.current_url, raw_href)
                        break
            except Exception:
                pass # Keep N/A if not found

            # ---------------------------------------------------------
            # IMAGES
            # ---------------------------------------------------------
            images = []
            try:
                # 1. Main Image
                img1_elem = None
                try:
                    img1_elem = driver.find_element(By.CSS_SELECTOR, "#scrolling_img img")
                except:
                    try:
                        img1_elem = driver.find_element(By.CSS_SELECTOR, "img.productListItemImg")
                    except: pass
                
                if img1_elem:
                    scroll_to_element(driver, img1_elem)
                    src = img1_elem.get_attribute("src")
                    if src: images.append(fix_url(driver.current_url, src))

                # 2. Slider Images
                slider = driver.find_elements(By.CLASS_NAME, "pdp-finishes-container")
                if slider:
                    scroll_to_element(driver, slider[0])
                    slides = driver.find_elements(By.CSS_SELECTOR, ".finishes-slider .slick-slide:not(.slick-cloned) img")
                    for img in slides:
                        # scroll_to_element(driver, img) # Hover logic often needed for lazy loading images
                        src = img.get_attribute("data-filename")
                        if not src: src = img.get_attribute("src")
                        if src:
                            full = fix_url(driver.current_url, src)
                            if full not in images: images.append(full)
            except: pass

            dim_img = "N/A"
            try:
                d_elem = driver.find_element(By.CSS_SELECTOR, "#collapse-dimensions img")
                scroll_to_element(driver, d_elem)
                dim_img = fix_url(driver.current_url, d_elem.get_attribute("src"))
            except: pass

            full_html = "N/A"
            try:
                full_html = driver.find_element(By.ID, "pdp-accordion").get_attribute("outerHTML")
            except: pass

            # Build Row
            row_data = {
                "Category": cat_name, 
                "Product URL": p_url,
                "Product Name": p_name,
                "SKU": sku,
                "Brand": "Tech Lighting",
                "Description": desc,
                "Lamping": lamping,
                "Full Description HTML": full_html,
                "Dimension Image": dim_img,
                "Spec Sheet": spec_sheet,
                "Image 1": images[0] if len(images) > 0 else "N/A",
                "Image 2": images[1] if len(images) > 1 else "N/A",
                "Image 3": images[2] if len(images) > 2 else "N/A",
                "Image 4": images[3] if len(images) > 3 else "N/A",
            }
            
            # SAVE TO CACHE
            product_data_cache[p_url] = row_data
            
            # ADD TO OUTPUT
            extracted_data.append(row_data)

            if (i + 1) % SAVE_BATCH_SIZE == 0:
                save_data(extracted_data, OUTPUT_FILE)

        except Exception as e:
            print(f"Error on {p_url}: {e}")

# ---------------------------------------------------------
# EXECUTION
# ---------------------------------------------------------
def signal_handler(sig, frame):
    print("\n[STOP] Saving data...")
    save_data(extracted_data, OUTPUT_FILE)
    sys.exit(0)

if __name__ == "__main__":
    signal.signal(signal.SIGINT, signal_handler)
    driver = setup_driver()
    
    try:
        tasks = collect_products(driver)
        if tasks:
            scrape_details(driver, tasks)
            save_data(extracted_data, OUTPUT_FILE)
            print("Done!")
        else:
            print("No products found.")
    finally:
        driver.quit()