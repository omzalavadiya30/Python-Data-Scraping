import time
import pandas as pd
import undetected_chromedriver as uc
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os # <-- Added for loading/saving files

# ------------------- User-Provided Category URLs -------------------
category_urls = [
    "https://www.arteriorshome.com/shop/new",
    "https://www.arteriorshome.com/shop/lighting/new-lighting",
    "https://www.arteriorshome.com/shop/furniture/new-furniture",
    "https://www.arteriorshome.com/shop/accessories/new-accessories",
    "https://www.arteriorshome.com/shop/wall/new-wall-decor",
    "https://www.arteriorshome.com/shop/furniture/dining/new-dining",
    "https://www.arteriorshome.com/shop/outdoor/new-outdoor",
    "https://www.arteriorshome.com/shop/outdoor/furniture",
    "https://www.arteriorshome.com/shop/outdoor/furniture/tables",
    "https://www.arteriorshome.com/shop/outdoor/furniture/seating",
    "https://www.arteriorshome.com/shop/outdoor/furniture/dining",
    "https://www.arteriorshome.com/shop/outdoor/furniture/lounge",
    "https://www.arteriorshome.com/shop/outdoor/lighting",
    "https://www.arteriorshome.com/shop/outdoor/rugs",
    "https://www.arteriorshome.com/shop/outdoor/accessories",
    "https://www.arteriorshome.com/shop/outdoor/outlet",
    "https://www.arteriorshome.com/shop/lighting/chandeliers",
    "https://www.arteriorshome.com/shop/lighting/chandeliers/decorative",
    "https://www.arteriorshome.com/shop/lighting/chandeliers/oversized",
    "https://www.arteriorshome.com/shop/lighting/chandeliers/linear",
    "https://www.arteriorshome.com/shop/lighting/chandeliers/natural",
    "https://www.arteriorshome.com/shop/lighting/pendants",
    "https://www.arteriorshome.com/shop/lighting/pendants/decorative",
    "https://www.arteriorshome.com/shop/lighting/pendants/oversized",
    "https://www.arteriorshome.com/shop/lighting/pendants/natural",
    "https://www.arteriorshome.com/shop/lighting/sconces",
    "https://www.arteriorshome.com/shop/lighting/sconces/decorative",
    "https://www.arteriorshome.com/shop/lighting/sconces/vanity",
    "https://www.arteriorshome.com/shop/lighting/sconces/task-sconce",
    "https://www.arteriorshome.com/shop/lighting/flush-mounts",
    "https://www.arteriorshome.com/shop/lighting/flush-mounts/decorative",
    "https://www.arteriorshome.com/shop/lighting/flush-mounts/oversized",
    "https://www.arteriorshome.com/shop/lighting/flush-mounts/semi-flush",
    "https://www.arteriorshome.com/shop/lighting/table-lamps",
    "https://www.arteriorshome.com/shop/lighting/table-lamps/decorative",
    "https://www.arteriorshome.com/shop/lighting/table-lamps/oversized",
    "https://www.arteriorshome.com/shop/lighting/table-lamps/task-lamp",
    "https://www.arteriorshome.com/shop/lighting/table-lamps/lamp-shades",
    "https://www.arteriorshome.com/shop/lighting/floor-lamps",
    "https://www.arteriorshome.com/shop/lighting/floor-lamps/decorative",
    "https://www.arteriorshome.com/shop/lighting/floor-lamps/task-floor-lamp",
    "https://www.arteriorshome.com/shop/lighting/floor-lamps/arc",
    "https://www.arteriorshome.com/shop/lighting/pipes-and-chains",
    "https://www.arteriorshome.com/shop/lighting/on-sale",
    "https://www.arteriorshome.com/shop/furniture/tables",
    "https://www.arteriorshome.com/shop/furniture/tables/coffee-and-cocktail-tables",
    "https://www.arteriorshome.com/shop/furniture/tables/accent-side-end-and-occassional-tables",
    "https://www.arteriorshome.com/shop/furniture/tables/dining-and-entry-tables",
    "https://www.arteriorshome.com/shop/furniture/tables/nightstands",
    "https://www.arteriorshome.com/shop/furniture/tables/desks",
    "https://www.arteriorshome.com/shop/furniture/seating",
    "https://www.arteriorshome.com/shop/furniture/seating/benches",
    "https://www.arteriorshome.com/shop/furniture/seating/ottomans-stools",
    "https://www.arteriorshome.com/shop/furniture/seating/bar-and-counter-stools",
    "https://www.arteriorshome.com/shop/furniture/seating/sofas-settees",
    "https://www.arteriorshome.com/shop/furniture/seating/swivel-chairs",
    "https://www.arteriorshome.com/shop/furniture/seating/lounge-chairs",
    "https://www.arteriorshome.com/shop/furniture/seating/arm-chairs",
    "https://www.arteriorshome.com/shop/furniture/seating/dining-chairs",
    "https://www.arteriorshome.com/shop/furniture/storage-shelving",
    "https://www.arteriorshome.com/shop/furniture/storage-shelving/cocktail-cabinets-bar-carts",
    "https://www.arteriorshome.com/shop/furniture/storage-shelving/cabinets",
    "https://www.arteriorshome.com/shop/furniture/storage-shelving/credenzas-consoles",
    "https://www.arteriorshome.com/shop/furniture/storage-shelving/bookshelves-etageres",
    "https://www.arteriorshome.com/shop/furniture/on-sale",
    "https://www.arteriorshome.com/shop/accessories/candles",
    "https://www.arteriorshome.com/shop/accessories/fireplace",
    "https://www.arteriorshome.com/shop/accessories/trays",
    "https://www.arteriorshome.com/shop/accessories/barware-and-entertaining",
    "https://www.arteriorshome.com/shop/accessories/objects-sculptures-and-bookends",
    "https://www.arteriorshome.com/shop/accessories/centerpieces-bowls",
    "https://www.arteriorshome.com/shop/accessories/decorative-boxes-containers",
    "https://www.arteriorshome.com/shop/accessories/vases-planters",
    "https://www.arteriorshome.com/shop/accessories/on-sale",
    "https://www.arteriorshome.com/shop/wall/decorative",
    "https://www.arteriorshome.com/shop/wall/mirrors",
    "https://www.arteriorshome.com/shop/wall/mirrors/full-length-mirrors",
    "https://www.arteriorshome.com/shop/wall/mirrors/vanity-mirrors",
    "https://www.arteriorshome.com/shop/wall/mirrors/mantel-mirrors",
    "https://www.arteriorshome.com/shop/wall/on-sale"
]

# ------------------- Config -------------------
OPERA_BINARY = r"C:\Users\HP\AppData\Local\Programs\Opera\opera.exe"
USER_DATA_DIR = r"C:\OperaProfileVPN"
CHROMEDRIVER_PATH = r"C:\WebDrivers\chromedriver-win64\chromedriver.exe"
CHROMIUM_MAJOR = 141  # Make sure this matches your Opera version base

# ------------------- Driver Setup -------------------
def setup_opera():
    options = Options()
    options.binary_location = OPERA_BINARY
    options.add_argument("--start-maximized")
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")
    options.add_argument(f"--user-data-dir={USER_DATA_DIR}")
    options.add_argument("--profile-directory=Default")
    options.add_argument("--disable-extensions")
    options.add_argument("--remote-debugging-port=0")

    service = Service(CHROMEDRIVER_PATH)
    print("Setting up Opera driver...")
    driver = uc.Chrome(service=service, options=options, version_main=CHROMIUM_MAJOR)
    driver.set_page_load_timeout(300) # Set page load timeout to 5 minutes
    return driver

def get_product_links_with_pagination(driver, category_url):
    """
    Visits a category page, scrapes all product links,
    and handles pagination to get all products in that category.
    """
    product_links = []
    
    # --- DEFINE YOUR SELECTORS HERE ---
    # This selector finds the product link on the grid
    product_selector = "div.kuName > a.klevuProductClick"
    # This selector finds the "Next" page button (link with text ">")
    next_page_xpath_selector = "//div[contains(@class, 'kuPagination')]//a[contains(@class, 'klevuPaginate') and normalize-space()='>']"
    # Selector to wait for the results to be present
    product_list_selector = "div.kuResults"
    
    try:
        driver.get(category_url)
        # Wait for the initial product list to load
        # Increased wait to 60s for slow-loading categories
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, product_list_selector))
        )
        time.sleep(2) # Extra time for JS to settle
    except TimeoutException:
        print(f"    !!! ERROR: Page {category_url} took too long to load products. Skipping category.")
        return [] # Return empty list if category times out
    except Exception as e:
        print(f"    Error loading page {category_url} or finding product list: {e}")
        return []

    page_count = 1
    while True:
        print(f"    ... Scraping page {page_count}")
        try:
            # Wait for product links to be present on the current page
            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, product_selector))
                )
            except TimeoutException:
                # It's possible a page has no products, check for pagination
                print("    ... No products found on this page.")
                pass # Will check for next page button

            # Find all product elements on the current page
            product_elements = driver.find_elements(By.CSS_SELECTOR, product_selector)
            
            new_products_found = 0
            for elem in product_elements:
                try:
                    prod_url = elem.get_attribute('href')
                    # Get name from h2 tag inside link, or title attribute as fallback
                    try:
                        prod_name = elem.find_element(By.CSS_SELECTOR, "h2.product-name").text.strip()
                    except:
                        prod_name = elem.get_attribute('title').strip()
                    
                    if not prod_name: # Fallback if h2 is empty and title is missing
                         prod_name = "Name not found"

                    if prod_url and prod_url not in [p['url'] for p in product_links]:
                        product_links.append({
                            'product_name': prod_name,
                            'url': prod_url
                        })
                        new_products_found += 1
                except Exception as e:
                    print(f"    ... Error scraping one product item: {e}")
            
            if product_elements:
                print(f"    ... Found {new_products_found} new products on this page. Total for category: {len(product_links)}")

            # Check for "Next" page button
            try:
                next_button = driver.find_element(By.XPATH, next_page_xpath_selector)
                
                print("    ... Clicking 'Next' page.")
                # We need to scroll to it and click, as it might be at the bottom
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
                time.sleep(0.5) # brief pause before click
                
                # Use JS click just in case it's obscured
                driver.execute_script("arguments[0].click();", next_button)
                
                page_count += 1
                
                # Wait for the page to update.
                print("    ... Waiting for page content to update...")
                # THIS IS THE FIX: Wait for the old "Next" button to go stale.
                # This proves the DOM has changed and the new page is loading.
                WebDriverWait(driver, 20).until(EC.staleness_of(next_button))
                
                # Add a small extra sleep for the new content to fully render
                time.sleep(2) 
                
            except NoSuchElementException:
                print("    ... No 'Next' (>) button found. Assuming end of category.")
                break # No more pages
            except TimeoutException:
                print("    ... Waited for page to update, but it didn't. Stopping category.")
                break # Page didn't update, stop.

        except Exception as e:
            print(f"    An error occurred during scraping: {e}")
            break
            
    return product_links

def save_data(data_list, filename):
    """Helper function to save all data to an Excel file."""
    if not data_list:
        print("No data to save.")
        return
        
    print(f"\nSaving {len(data_list)} products to {filename}...")
    try:
        df = pd.DataFrame(data_list)
        df.to_excel(filename, index=False)
        print("Save successful.")
    except Exception as e:
        print(f"!!! Error saving file: {e}")

def scrape_product_details(driver, product_url):
    """
    Visits a single product page and scrapes all required details.
    Includes logic to handle mid-scrape Cloudflare challenges.
    """
    
    # --- NEW: Retry loop to handle CF and load errors ---
    # Try to load the page up to 2 times
    for attempt in range(2):
        try:
            driver.get(product_url)
            
            # --- NEW: Cloudflare Check ---
            # Immediately check if we are on a challenge page (short timeout)
            try:
                # Look for a common Cloudflare element (e.g., iframe)
                WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'iframe[title*="Cloudflare"], iframe[title*="challenge"]'))
                )
                
                # If found, print and wait for it to be solved (long timeout)
                print(f" ----------------------------------------------------")
                print(f"    !!! Cloudflare challenge detected on attempt {attempt+1} for {product_url}")
                print(f"    ... Waiting for solver (max 60s)...")
                
                # Now, wait for the *real* page content to appear,
                # which signals the challenge is passed.
                WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "h1.page-title span.base"))
                )
                print(f"    ... Cloudflare challenge passed!")
                print(f" ----------------------------------------------------")
                
            except TimeoutException:
                # This is the GOOD path: No Cloudflare iframe was found in 3 seconds.
                # Now we just do the *normal* wait for the page title.
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "h1.page-title span.base"))
                )
            
            # --- If we get here, the page is loaded (either directly or after solving CF) ---
            break # Exit the retry loop and proceed to scraping

        except TimeoutException:
            # This exception means one of the *main* waits failed 
            # (either the 60s post-CF wait, or the 30s normal wait).
            print(f"    !!! ERROR: Page {product_url} timed out on attempt {attempt+1}. Retrying...")
            time.sleep(3) # Wait a bit before retrying
            
        except Exception as e:
            # Other navigation error
            print(f"    !!! ERROR: Navigating to {product_url} failed on attempt {attempt+1}: {e}. Retrying...")
            time.sleep(3)
    
    else: 
        # This 'else' block runs if the 'for' loop completes without a 'break'
        print(f" ----------------------------------------------------")
        print(f"    !!! CRITICAL: Failed to load {product_url} after 2 attempts. Skipping product.")
        print(f" ----------------------------------------------------")
        return None # Signal failure

    # --- END OF NEW LOGIC ---
    # The rest of the function is your original, working scraping logic.
    # It only runs if the loop above 'break's successfully.

    details = {
        'Category': '',
        'Product_URL': product_url,
        'Product Name': '', 
        'SKU': '',
        'Brand': 'Arteriors Home',
        'Attr_Width_In': '',
        'Attr_Depth_In': '',
        'Attr_Height_In': '',
        'Attr_Diameter_In': '',
        'Attr_Width_Cm': '',
        'Attr_Depth_Cm': '',
        'Attr_Height_Cm': '',
        'Attr_Diameter_Cm': '',
        'Description': '',
        'Full_Description_HTML': '',
        'Tearsheet': '',
        'SpecSheet': '',
        'Assembly_Instruction': '',
    }

    # 1. Category (Breadcrumbs)
    try:
        crumbs = driver.find_elements(By.CSS_SELECTOR, "ul.breadcrumbs li")
        # Get text, strip whitespace, filter out empty strings
        crumb_text = [c.text.strip() for c in crumbs if c.text.strip()]
        details['Category'] = " > ".join(crumb_text)
    except Exception:
        details['Category'] = None

    # 2. Product Name
    try:
        details['Product Name'] = driver.find_element(By.CSS_SELECTOR, "h1.page-title span.base").text.strip()
    except Exception:
        details['Product Name'] = None
        
    # 3. SKU
    try:
        details['SKU'] = driver.find_element(By.CSS_SELECTOR, "div.product-sku h2.pro-sku").text.strip()
    except Exception:
        details['SKU'] = None
        
    # 4. Dimensions (in/cm) - using data attributes (no click needed)
    dim_map = {
        'Width': ('Attr_Width_In', 'Attr_Width_Cm'),
        'Depth': ('Attr_Depth_In', 'Attr_Depth_Cm'),
        'Height': ('Attr_Height_In', 'Attr_Height_Cm'),
        'Diameter': ('Attr_Diameter_In', 'Attr_Diameter_Cm'),
    }
    # Initialize all keys to None
    for in_key, cm_key in dim_map.values():
        details[in_key] = None
        details[cm_key] = None
        
    try:
        dim_spans = driver.find_elements(By.CSS_SELECTOR, "div#dimensions span.product-metric")
        for span in dim_spans:
            label = span.get_attribute('data-label')
            in_val = span.get_attribute('data-in')
            cm_val = span.get_attribute('data-cm')
            
            for key_name, (in_key, cm_key) in dim_map.items():
                if key_name in label:
                    details[in_key] = in_val
                    details[cm_key] = cm_val
                    break # Move to next span
    except Exception as e:
        print(f"       - Warning: Could not scrape Dimensions: {e}")

    # 5. Description
    try:
        details['Description'] = driver.find_element(By.CSS_SELECTOR, "div.product.attribute.overview > div.value").text.strip()
    except Exception:
        details['Description'] = None
        
    # 6. Full Description (HTML)
    try:
        details['Full_Description_HTML'] = driver.find_element(By.CSS_SELECTOR, "div.product-info-main").get_attribute('outerHTML')
    except Exception:
        details['Full_Description_HTML'] = None
        
    # 7. Technical Documents
    details['Tearsheet'] = None
    details['SpecSheet'] = None
    details['Assembly_Instruction'] = None
    try:
        # Find the accordion header
        accordion_header = driver.find_element(By.XPATH, "//h3[normalize-space()='Technical Documents']")
        # Find its parent 'holder' div
        parent_holder = accordion_header.find_element(By.XPATH, "..")
        
        # Click it if it's not already open
        if "open" not in parent_holder.get_attribute("class"):
            driver.execute_script("arguments[0].click();", accordion_header)
            time.sleep(1) # Wait for accordion to open

        # Find doc links within the now-open accordion
        doc_links = parent_holder.find_elements(By.CSS_SELECTOR, "a.tearsheet")
        for link in doc_links:
            if "tear-report" in link.get_attribute("class"):
                details['Tearsheet'] = link.get_attribute("data-id")
            elif "spec-report" in link.get_attribute("class"):
                details['SpecSheet'] = link.get_attribute("data-id")
            elif "ai-report" in link.get_attribute("class"):
                details['Assembly_Instruction'] = link.get_attribute("data-id")
    except Exception:
        # This is not an error, some products just don't have docs
        pass 

    # ------------------- NEW FIX: SCROLL INTO VIEW -------------------
    # We must scroll the product info area into view
    # so that the image thumbnails at the bottom become "visible" to Selenium.
    try:
        main_info = driver.find_element(By.CSS_SELECTOR, "div.product-info-main")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", main_info)
        time.sleep(1) # Wait for scroll to finish
    except NoSuchElementException:
        print("       - Warning: Could not find main product block to scroll to.")
        pass # Continue anyway
    # ----------------- END OF NEW FIX -----------------

    # # 8. Images
    # # Initialize all image keys to None first
    # for i in range(4):
    #     details[f'Image{i+1}'] = None
        
    # image_links = []
    # try:
    #     # --- STRATEGY 1: Wait for a VISIBLE slide in 'div#more-views' ---
    #     selector_A = "div#more-views .slick-track .slick-slide:not(.slick-cloned) a[data-image]"
    #     try:
    #         # Wait for the element to be VISIBLE (i.e., on-screen)
    #         WebDriverWait(driver, 20).until(
    #             EC.visibility_of_element_located((By.CSS_SELECTOR, selector_A))
    #         )
    #         image_elements = driver.find_elements(By.CSS_SELECTOR, selector_A)
    #         image_links = [el.get_attribute('data-image') for el in image_elements if el.get_attribute('data-image')]
        
    #     except TimeoutException:
    #         # --- STRATEGY 2: If A fails, wait for a VISIBLE slide in 'div.product.media' ---
    #         selector_B = "div.product.media div.slick-track .slick-slide:not(.slick-cloned) a[data-image]"
    #         try:
    #             # Also wait for VISIBILITY here
    #             WebDriverWait(driver, 10).until(
    #                 EC.visibility_of_element_located((By.CSS_SELECTOR, selector_B))
    #             )
    #             image_elements = driver.find_elements(By.CSS_SELECTOR, selector_B)
    #             image_links = [el.get_attribute('data-image') for el in image_elements if el.get_attribute('data-image')]
            
    #         except TimeoutException:
    #             # --- BOTH FAILED ---
    #             pass

    #     # --- Now, process the results ---
    #     if not image_links:
    #         # This warning will ONLY print if both Strategy A and Strategy B timed out.
    #         print(f"       - Warning: Waited for images, but none were found for this product.")
            
    #     else:
    #         # De-duplicate the list while preserving order
    #         unique_image_links = []
    #         for link in image_links:
    #             if link not in unique_image_links:
    #                 unique_image_links.append(link)
            
    #         for i in range(4): # Get up to 4 images
    #             if i < len(unique_image_links):
    #                 details[f'Image{i+1}'] = unique_image_links[i]

    # except Exception as e:
    #     print(f"       - Warning: An error occurred during image scraping: {e}")

    # ------------------- UNIVERSAL IMAGE SCRAPING (FINAL VERSION) -------------------
    # Initialize all image keys to None
    for i in range(4):
        details[f'Image{i+1}'] = None

    try:
        # Scroll entire page to trigger lazy loading
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1.5)
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(1)

        # Collect ALL image URLs on page
        all_imgs = driver.find_elements(By.TAG_NAME, "img")

        raw_links = []
        for img in all_imgs:
            src = img.get_attribute("src")
            data_src = img.get_attribute("data-src")
            data_image = img.get_attribute("data-image")

            for link in [src, data_src, data_image]:
                if link and link.startswith("http"):
                    raw_links.append(link)

        # Optional: Filter only product media folder
        filtered = [x for x in raw_links if "/media/catalog/" in x or "/Arteriors/" in x]

        # Fallback: If filter finds nothing, use raw list
        final_links = filtered if filtered else raw_links

        # Deduplicate
        unique = []
        for u in final_links:
            if u not in unique:
                unique.append(u)

        # Save top 4 images
        for i in range(min(4, len(unique))):
            details[f'Image{i+1}'] = unique[i]

    except Exception as e:
        print(f"       - Warning: Universal image scraping error: {e}")
            
    return details

def parse_category_name(url):
    """Helper function to create a clean category name from the URL."""
    try:
        return url.split("/shop/")[1].replace("/", " > ")
    except:
        return url

# ------------------- Main Scraping Logic -------------------
if __name__ == "__main__":
    
    output_file = 'arteriors_products_output.xlsx'
    scraped_product_urls = set()
    all_data = []

    # Load existing data to avoid re-scraping
    if os.path.exists(output_file):
        try:
            print(f"Loading existing data from {output_file}...")
            df = pd.read_excel(output_file)
            # Ensure 'Product_URL' column exists before proceeding
            if 'Product_URL' in df.columns:
                all_data = df.to_dict('records')
                # Add valid URLs to the set
                scraped_product_urls = set(df['Product_URL'].dropna())
                print(f"Loaded {len(all_data)} existing products. {len(scraped_product_urls)} URLs already scraped.")
            else:
                print(f"'{output_file}' found, but 'Product_URL' column is missing. Starting fresh.")
                all_data = []
                scraped_product_urls = set()
        except Exception as e:
            print(f"Error loading {output_file}: {e}. Starting fresh.")
            all_data = []
            scraped_product_urls = set()
    
    driver = setup_opera()
    total_products_scraped_session = 0

    try:
        # ------------------------------------------------------------------
        # --- ATTEMPT AUTOMATIC CLOUDFLARE VERIFICATION ---
        # ------------------------------------------------------------------
        print("\n" + "="*60)
        print("      Attempting to bypass Cloudflare...")
        print("      This may take up to 60 seconds.")
        print("="*60)
        
        # Load the first page to trigger the challenge
        driver.get(category_urls[0])
        
        # Wait up to 60 seconds for the product list to appear.
        # This gives undetected-chromedriver time to solve the challenge.
        try:
            print("... Waiting for page to load (max 60 seconds)...")
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.kuResults"))
            )
            print("... Cloudflare bypassed successfully! Resuming script...")
        except TimeoutException:
            print("\n" + "!"*60)
            print("    AUTOMATIC BYPASS FAILED. Cloudflare is still blocking.")
            print("    The script cannot continue.")
            print("    Try running the script again. Sometimes it works on the 2nd/3rd try.")
            print("    If it keeps failing, the manual 'input()' step from the")
            print("    previous version may be required.")
            print("!"*60)
            raise # Stop the script
        # ------------------------------------------------------------------
        
        print("\nCloudflare verification assumed complete. Resuming script...")
        
        total_categories = len(category_urls)
        
        # Loop through all provided category URLs
        for i, category_url in enumerate(category_urls):
            
            products_list = []
            # Get all product links for this category
            if i == 0 and driver.current_url.startswith(category_urls[0]):
                print(f"\nScraping Category {i+1}/{total_categories}: {category_url} (already open)")
                # We are already on this page from the Cloudflare check
                products_list = get_product_links_with_pagination(driver, driver.current_url)
            else:
                print(f"\nScraping Category {i+1}/{total_categories}: {category_url}")
                products_list = get_product_links_with_pagination(driver, category_url)
            
            if not products_list:
                print(f"--- No products found for category {i+1}. Moving to next. ---")
                continue

            # Filter out products we've already scraped
            products_to_scrape = [p for p in products_list if p['url'] not in scraped_product_urls]
            
            if not products_to_scrape:
                print(f"--- All {len(products_list)} products in this category already scraped. Skipping. ---")
                continue
                
            print(f"--- Found {len(products_to_scrape)} new products to scrape in this category (out of {len(products_list)} total). ---")

            # Now, scrape details for each product in this category
            for j, product_summary in enumerate(products_to_scrape):
                # NEW LOGGING
                print(f"  Scraping product {j+1}/{len(products_to_scrape)} -> {product_summary['product_name']} -> {product_summary['url']}")
                
                try:
                    details = scrape_product_details(driver, product_summary['url'])
                    
                    if details: # If scraping wasn't skipped due to timeout
                        all_data.append(details)
                        scraped_product_urls.add(product_summary['url']) # Add to set to avoid re-scraping
                        total_products_scraped_session += 1
                    
                    # PERIODIC SAVE
                    if total_products_scraped_session > 0 and total_products_scraped_session % 50 == 0:
                        print(f"\n--- Auto-saving progress: {total_products_scraped_session} new products scraped this session. ---")
                        save_data(all_data, output_file)
                        
                except Exception as e:
                    print(f"    !!! CRITICAL ERROR scraping {product_summary['url']}: {e}")
                    print("    ... skipping this product.")

            print(f"--- Finished category {i+1}. Total products in database: {len(all_data)} ---")
            # Be respectful, add a delay between categories
            time.sleep(2)

    except KeyboardInterrupt:
        print("\n" + "!"*60)
        print("    Ctrl+C detected! User interrupted script. Saving backup...")
        print("!"*60)
    
    except Exception as e:
        print(f"\nA critical error occurred in the main loop: {e}")
    
    finally:
        print("\nScraping finished or interrupted. Saving final data...")
        if all_data:
            save_data(all_data, output_file)
        print("Closing browser.")
        driver.quit()
    
    # ------------------- Final Save -------------------
    if all_data:
        print(f"\nFinal check: Successfully saved {len(all_data)} items to '{output_file}'")
    else:
        print("\nNo data was scraped. The DataFrame is empty.")
