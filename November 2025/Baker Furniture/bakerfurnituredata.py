import pandas as pd
import time
from urllib.parse import urljoin
import re  # For parsing dimensions
import os  # For final file cleanup

# Selenium imports
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, InvalidSessionIdException
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

# --- Dictionary to map table abbreviations to Excel column names ---
DIMENSION_MAP = {
    "W": ("Attr_Width_In", "Attr_Width_Cm"),
    "D": ("Attr_Depth_In", "Attr_Depth_Cm"),
    "H": ("Attr_Height_In", "Attr_Height_Cm"),
    "DI": ("Attr_Depth_Inside_In", "Attr_Depth_Inside_Cm"),
    "SH": ("Attr_Seat_Height_In", "Attr_Seat_Height_Cm"),
    "SD": ("Attr_Seat_Depth_In", "Attr_Seat_Depth_Cm"),
    "AH": ("Attr_Arm_Height_In", "Attr_Arm_Height_Cm"),
    "WI": ("Attr_Width_Inside_In", "Attr_Width_Inside_Cm"),
    "DIA": ("Attr_Diameter_In", "Attr_Diameter_Cm"),  # <-- ADDED 'DIA'
    "COM": ("Attr_COM_Yd", "Attr_COM_M"),
    "COL": ("Attr_COL_Sqft", "Attr_COL_M")
}

# --- *** NEW: Create a sorted list of keys (longest first) *** ---
# This is the fix. It checks "DI" before "D", "SH" before "H", etc.
SORTED_DIMENSION_KEYS = sorted(DIMENSION_MAP.keys(), key=len, reverse=True)


def initialize_driver():
    """
    Creates and returns a new, robust Selenium WebDriver instance.
    """
    print("  -> Initializing new Chrome driver...")
    s = Service(ChromeDriverManager().install())
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36"
    )
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("log-level=3")

    driver = webdriver.Chrome(service=s, options=options)
    print("  -> New driver is ready.")
    return driver


def get_category_links(driver, wait, base_url):
    """
    Part 1: Gets all unique category URLs to visit.
    This version skips 'View All' parent links if child links exist.
    """

    print("--- PART 1: Scraping Category URLs from Navbar ---")
    exclude_links_by_text = {
        "Living": {"Bespoke Custom Seating", "Bespoke in Motion", "Bespoke Custom Pillows"},
        "Dining": {"Bespoke Custom Pillows"},
        "Bedroom": {"Bespoke Custom Beds", "Bespoke Custom Pillows"},
        "Workspace": {"Bespoke Custom Pillows"},
        "Custom": {"Custom"},  # Skips the main /custom/ page
        "Outdoor": set(),
        "Textiles": set()
    }
    target_categories = ["Living", "Dining", "Bedroom", "Workspace", "Outdoor", "Textiles", "Custom"]
    all_category_links = []

    driver.get(base_url)
    try:
        wait.until(EC.element_to_be_clickable((By.ID, "onetrust-accept-btn-handler"))).click()
        print("Closed cookie banner.")
    except Exception:
        print("No cookie banner found or couldn't close it.")

    for category_name in target_categories:
        print(f"Finding sub-categories for: {category_name}")
        try:
            main_link_xpath = f"//a[contains(@class, 'primaryNav-link') and @data-menu-name='{category_name}']"
            main_link = wait.until(EC.element_to_be_clickable((By.XPATH, main_link_xpath)))
            main_link.click()

            submenu_xpath = (
                f"//div[contains(@class, 'primaryNav-submenu') and "
                f"@data-menu-name='{category_name}' and contains(@class, 'active')]"
            )
            wait.until(EC.visibility_of_element_located((By.XPATH, submenu_xpath)))

            # --- Get CHILD links first ---
            sub_links_xpath = f"{submenu_xpath}//a[contains(@class, 'subNav-link')]"
            sub_links = driver.find_elements(By.XPATH, sub_links_xpath)

            # If no child links were found, get the parent links
            if not sub_links:
                sub_links_xpath = f"{submenu_xpath}//a[contains(@class, 'subNav-link_main')]"
                sub_links = driver.find_elements(By.XPATH, sub_links_xpath)

            links_found_count = 0
            for link in sub_links:
                link_text = link.text.strip()
                link_url = link.get_attribute('href')

                if link_text in exclude_links_by_text.get(category_name, set()):
                    print(f"  -> Skipping (duplicate/excluded): {link_text}")
                    continue

                full_url = urljoin(base_url, link_url)

                if link_text and link_url:
                    all_category_links.append({
                        "main_category": category_name,
                        "sub_category_name": link_text,
                        "category_url": full_url
                    })
                    links_found_count += 1

            print(f"Found and added {links_found_count} sub-category links for {category_name}.")

        except Exception as e:
            print(f"Error processing category '{category_name}': {e}")
            driver.find_element(By.TAG_NAME, 'body').click()
            time.sleep(0.5)

    print("\n--- De-duplicating category URLs ---")
    if not all_category_links:
        raise Exception("No category links were found. Exiting.")

    links_df = pd.DataFrame(all_category_links)
    links_df.drop_duplicates(subset=["category_url"], keep="first", inplace=True)
    category_links_to_visit = links_df.to_dict('records')

    print(f"Collected {len(category_links_to_visit)} UNIQUE category pages to visit.\n")
    return category_links_to_visit


def get_products_and_category(driver, wait, category_url):
    """
    Part 2: Visits a single category page, scrapes its product URLs
    AND its category breadcrumb.
    Returns: (list_of_urls, category_string)
    """
    product_urls = set()
    category_string = "N/A"  # Default to N/A

    try:
        driver.get(category_url)
        # Wait for either type of product list to appear
        wait.until(EC.any_of(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "ul#js-listing-quickLookView")),
            EC.visibility_of_element_located((By.CSS_SELECTOR, "ul.dropBlocks"))
        ))
        time.sleep(2)

        # --- SCRAPE CATEGORY FROM THIS PAGE ---
        try:
            breadcrumbs = driver.find_elements(By.CSS_SELECTOR, "ol.wayfinding li a")
            crumb_texts = [b.text.strip() for b in breadcrumbs]  # Get ALL crumbs
            category_string = " > ".join(crumb_texts).upper()
            print(f"  -> Category found: {category_string}")
        except Exception as e:
            print(f"  -> Warning: Could not find category breadcrumb. Defaulting to N/A. {e}")

        # Method A: "View More" Button Clicker (Click ONCE)
        try:
            view_more_button = driver.find_element(By.CSS_SELECTOR, "button[data-disclosure-role='trigger']")
            if view_more_button.is_displayed():
                print("  -> Clicking 'View More' button ONCE...")
                driver.execute_script("arguments[0].click();", view_more_button)
                time.sleep(3)
        except Exception:
            pass  # No button is not an error

        # Method B: "Infinite Scroll"
        try:
            last_height = driver.execute_script("return document.body.scrollHeight")
            while True:
                footer_element = driver.find_element(By.CSS_SELECTOR, "div.footer-subscribe")
                footer_location = footer_element.location['y']
                scroll_target = max(0, footer_location - 300)
                driver.execute_script(f"window.scrollTo(0, {scroll_target});")
                time.sleep(4)
                new_height = driver.execute_script("return document.body.scrollHeight")
                if new_height == last_height:
                    break
                last_height = new_height
        except NoSuchElementException:
            # Fallback scroll
            body = driver.find_element(By.TAG_NAME, 'body')
            last_height = driver.execute_script("return document.body.scrollHeight")
            for _ in range(5):
                body.send_keys(Keys.PAGE_DOWN)
                time.sleep(3)
                new_height = driver.execute_script("return document.body.scrollHeight")
                if new_height == last_height:
                    break
                last_height = new_height

        # --- Find all product links ---
        product_links = driver.find_elements(
            By.CSS_SELECTOR,
            "ul.dropBlocks div.feature-label > a, li[data-quick-look-role='item'] div.feature-label > a"
        )

        for link_element in product_links:
            try:
                product_url = link_element.get_attribute('href')
                if product_url:
                    product_urls.add(product_url)
            except Exception:
                pass  # Ignore stale links

    except InvalidSessionIdException as e:
        raise e
    except Exception as e:
        print(f"  -> ERROR: Failed to scrape page {category_url}. Error: {e}")

    return list(product_urls), category_string


def extract_largest_scene7_url(srcset_value):
    """
    Takes a srcset string and returns ONLY the largest Scene7 image URL.
    Fix: split on ', ' (comma+space) so commas inside URLs don't break it.
    """
    if not srcset_value:
        return "N/A"
    
    # Split ONLY on ', ' (or comma + whitespace), NOT every comma
    parts = [p.strip() for p in re.split(r',\s+', srcset_value) if p.strip()]
    if not parts:
        return "N/A"

    # Last candidate = largest resolution
    last = parts[-1]
    url = last.split()[0].strip()   # take everything before "1556w"/"778w"
    return url


def scrape_product_details(driver, wait, product_url, category_string):
    """
    Part 3: Visits a single product page and scrapes all required details.
    """

    # Initialize all data points to "N/A" as requested
    product_data = {
        "Category": category_string,  # Use the category passed from Part 2
        "Product URL": product_url,
        "Product Name": "N/A",
        "SKU": "N/A",
        "Brand": "Baker Furniture",
        "Description": "N/A",
        "Full Description (HTML)": "N/A",
        "Image1": "N/A",
        "Image2": "N/A",
        "Image3": "N/A",
        "Image4": "N/A",
    }

    # Initialize all dimension fields
    for abbr, (in_col, cm_col) in DIMENSION_MAP.items():
        product_data[in_col] = "N/A"
        product_data[cm_col] = "N/A"

    try:
        driver.get(product_url)

        # Wait for product name to be visible
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "h1[itemprop='name']")))

        # --- Scrape Name ---
        try:
            product_data["Product Name"] = driver.find_element(
                By.CSS_SELECTOR, "h1[itemprop='name']"
            ).text.strip()
        except Exception:
            print(f"     - Warning: Could not find Product Name.")

        # --- Scrape SKU ---
        try:
            product_data["SKU"] = driver.find_element(
                By.CSS_SELECTOR, "div.feature-meta_relaxed span[itemprop='sku']"
            ).text.strip()
        except Exception:
            print(f"     - Warning: Could not find SKU.")

        # --- Scrape Description ---
        try:
            desc_container = driver.find_element(
                By.XPATH, "//h4[contains(text(), 'About This Product')]/ancestor::div[@class='wrapper']"
            )
            all_text_nodes = driver.execute_script(
                "return Array.from(arguments[0].childNodes)"
                ".filter(node => node.nodeType === Node.TEXT_NODE)"
                ".map(node => node.textContent.trim())"
                ".filter(text => text.length > 0)",
                desc_container
            )
            product_data["Description"] = " ".join(all_text_nodes).strip()
        except Exception:
            print(f"     - Warning: Could not find Description.")

        # --- Scrape Full Description (HTML) ---
        try:
            full_desc_element = driver.find_element(
                By.XPATH, "//h2[contains(text(), 'Detailed Dimensions')]/ancestor::div[contains(@class, 'grid')]"
            )
            product_data["Full Description (HTML)"] = full_desc_element.get_attribute('outerHTML')
        except Exception:
            print(f"     - Warning: Could not find Full Description HTML.")

        # --- Scrape IMAGES (FIXED) ---
        try:
            # IMAGE 1: primary image from the main imgSwitcher-primary-img srcset
            try:
                primary_img = driver.find_element(
                    By.CSS_SELECTOR,
                    "img.imgSwitcher-primary-img[data-image-switcher-role='primary-image']"
                )
                srcset_primary = primary_img.get_attribute("srcset")
                product_data["Image1"] = extract_largest_scene7_url(srcset_primary)
            except Exception as e:
                print(f"     - Warning: Could not get primary image: {e}")

            # IMAGE 2–4: from thumbnail buttons' data-image-switcher-srcset
            try:
                thumbs = driver.find_elements(
                    By.CSS_SELECTOR,
                    "ol.imgSwitcher-nav button[data-image-switcher-srcset]"
                )
                for i, t in enumerate(thumbs[:3]):  # next 3 images only
                    srcset_thumb = t.get_attribute("data-image-switcher-srcset")
                    product_data[f"Image{i+2}"] = extract_largest_scene7_url(srcset_thumb)
            except Exception as e:
                print(f"     - Warning: Could not get thumbnail images: {e}")
        except Exception as e:
            print(f"     - Warning: Could not find Images. {e}")

        # --- Scroll to Dimensions ---
        try:
            scroll_target = driver.find_element(By.CSS_SELECTOR, "ol.imgSwitcher-nav")
            driver.execute_script("arguments[0].scrollIntoView(true);", scroll_target)
            print("     - Scrolled to dimensions area.")
            time.sleep(1)
        except Exception:
            print(f"     - Warning: Could not find 'imgSwitcher-nav' to scroll to. Dimensions might be missing.")

        # --- DIMENSION LOGIC ---
        try:
            dim_headers = driver.find_elements(
                By.XPATH,
                "//div[contains(@class, 'feature-bd')]//table//th[@scope='row']/abbr"
            )

            for header in dim_headers:
                abbr_text = header.text.strip()
                abbr_key = None

                for key in SORTED_DIMENSION_KEYS:
                    if abbr_text.upper() == key.upper():
                        abbr_key = key
                        break

                if abbr_key:
                    us_col, metric_col = DIMENSION_MAP[abbr_key]
                    row = header.find_element(By.XPATH, "./ancestor::tr")
                    values = row.find_elements(By.TAG_NAME, "td")

                    if len(values) == 2:
                        val1_text = values[0].text.strip()
                        val2_text = values[1].text.strip()
                        val1_num = re.sub(r"[^\d\.]", "", val1_text)
                        val2_num = re.sub(r"[^\d\.]", "", val2_text)

                        us_unit_key = us_col.split('_')[-1].lower()
                        if us_unit_key == "sqft":
                            us_unit_key = "sq ft"
                        metric_unit_key = metric_col.split('_')[-1].lower()

                        if us_unit_key in val1_text.lower():
                            product_data[us_col] = val1_num
                        elif metric_unit_key in val1_text.lower():
                            product_data[metric_col] = val1_num

                        if us_unit_key in val2_text.lower():
                            product_data[us_col] = val2_num
                        elif metric_unit_key in val2_text.lower():
                            product_data[metric_col] = val2_num

        except Exception as e:
            print(f"     - Warning: Error parsing dimensions. {e}")

    except InvalidSessionIdException as e:
        print(f"     - FATAL ERROR: Session ID is invalid. Browser has crashed.")
        raise e
    except Exception as e:
        print(f"     - FATAL ERROR scraping page {product_url}. Error: {e}")
        return product_data

    return product_data


# --- Main Script Execution ---

def main():
    print("Setting up WebDriver...")
    driver = initialize_driver()
    wait = WebDriverWait(driver, 20)
    base_url = "https://www.bakerfurniture.com/"

    output_file = "baker_furniture_product_DETAILS.xlsx"
    checkpoint_file = "baker_furniture_checkpoint.xlsx"
    failed_file = "baker_furniture_FAILED_PRODUCTS.xlsx"

    all_scraped_data = []
    all_failed_items = []
    known_product_urls = set()  # Master set to de-duplicate products across categories

    # Clean up old files for a fresh run
    if os.path.exists(output_file):
        os.remove(output_file)
    if os.path.exists(checkpoint_file):
        os.remove(checkpoint_file)
    if os.path.exists(failed_file):
        os.remove(failed_file)

    try:
        # --- Part 1: Get all category URLs ---
        category_links_to_visit = get_category_links(driver, wait, base_url)

        if not category_links_to_visit:
            print("No category URLs were found. Exiting.")
            return

        print(f"\n--- STARTING MAIN SCRAPE: {len(category_links_to_visit)} CATEGORIES ---")

        # --- Loop 1: By Category ---
        for i, cat_info in enumerate(category_links_to_visit):
            print(f"\n--- Processing Category {i+1}/{len(category_links_to_visit)}: {cat_info['sub_category_name']} ---")

            try:
                product_urls_on_page, category_string = get_products_and_category(
                    driver, wait, cat_info['category_url']
                )

            except InvalidSessionIdException:
                print("  -> BROWSER CRASHED (InvalidSessionIdException) during Part 2. Restarting driver...")
                try:
                    driver.quit()
                except Exception:
                    pass
                driver = initialize_driver()
                wait = WebDriverWait(driver, 20)

                print("  -> Retrying last category...")
                try:
                    product_urls_on_page, category_string = get_products_and_category(
                        driver, wait, cat_info['category_url']
                    )
                except Exception as e:
                    print(f"  -> FATAL: Retry failed for category {cat_info['category_url']}. Skipping. Error: {e}")
                    all_failed_items.append({
                        "url": cat_info['category_url'],
                        "error": f"Category page scrape failed: {e}"
                    })
                    continue

            print(f"  -> Found {len(product_urls_on_page)} products for this category. Now scraping details...")

            # --- De-duplicate product URLs against our master list ---
            new_products_to_scrape = []
            for url in product_urls_on_page:
                if url not in known_product_urls:
                    new_products_to_scrape.append(url)
                    known_product_urls.add(url)

            if not new_products_to_scrape:
                print("  -> All products in this category have already been scraped. Skipping.")
                continue

            print(f"  -> {len(new_products_to_scrape)} are new. Total unique products so far: {len(known_product_urls)}")

            # --- Loop 2: By Product ---
            for j, url in enumerate(new_products_to_scrape):

                product_name_for_log = "..."

                try:
                    product_details = scrape_product_details(driver, wait, url, category_string)

                    if product_details.get("Product Name") != "N/A":
                        product_name_for_log = product_details['Product Name']

                except InvalidSessionIdException:
                    print(f"    Scraping ({j+1}/{len(new_products_to_scrape)}): {product_name_for_log} -> {url}")
                    print("      -> BROWSER CRASHED (InvalidSessionIdException). Restarting driver...")
                    try:
                        driver.quit()
                    except Exception:
                        pass

                    driver = initialize_driver()
                    wait = WebDriverWait(driver, 20)

                    print("      -> Retrying last product URL...")
                    try:
                        product_details = scrape_product_details(driver, wait, url, category_string)
                    except Exception as e:
                        print(f"      -> FATAL: Retry failed for {url}. Skipping. Error: {e}")
                        all_failed_items.append({"url": url, "error": str(e)})
                        continue

                except Exception as e:
                    print(f"    Scraping ({j+1}/{len(new_products_to_scrape)}): {product_name_for_log} -> {url}")
                    print(f"      -> FATAL: An unknown error occurred for {url}. Skipping. Error: {e}")
                    all_failed_items.append({"url": url, "error": str(e)})
                    continue

                print(f"    Scraping ({j+1}/{len(new_products_to_scrape)}): {product_details.get('Product Name', '...')} -> {url}")

                if product_details.get("Product Name") != "N/A":
                    print(f"      -> Successfully scraped.")
                else:
                    print(f"      -> WARNING: Scraped page but no product name found for {url}.")
                    all_failed_items.append({"url": url, "error": "No product name found."})

                all_scraped_data.append(product_details)

            # --- Checkpoint Save (After EACH category) ---
            if all_scraped_data:
                print(f"\n--- Checkpoint: Saving all {len(all_scraped_data)} products collected so far... ---")
                df_check = pd.DataFrame(all_scraped_data)
                df_check.to_excel(checkpoint_file, index=False)
                print("--- Checkpoint save complete. ---\n")

    except KeyboardInterrupt:
        print("\n--- KeyboardInterrupt (Ctrl+C) detected! ---")
        print("Stopping scraper and saving all data collected so far...")

    except Exception as e:
        print(f"\n--- An unexpected error occurred in the main loop: {e} ---")
        print("Saving all data collected so far...")

    finally:
        # --- Final Save ---
        if all_scraped_data:
            print(f"\n--- Saving all {len(all_scraped_data)} collected products to final file... ---")
            df = pd.DataFrame(all_scraped_data)
            df.drop_duplicates(subset=["Product URL"], keep="last", inplace=True)
            df.to_excel(output_file, index=False)
            print(f"Successfully saved {len(df)} unique products to {output_file}")

            if os.path.exists(checkpoint_file):
                os.remove(checkpoint_file)
                print("--- Checkpoint file removed. ---")
        else:
            print("No product data was collected to save.")

        # --- Save Failed URLs (if any) ---
        if all_failed_items:
            print(f"\n--- WARNING: {len(all_failed_items)} products/categories failed to scrape. ---")
            df_failed = pd.DataFrame(all_failed_items)
            df_failed.to_excel(failed_file, index=False)
            print(f"--- Saved a list of failed items to {failed_file} ---")

        print("Cleaning up and closing WebDriver.")
        if 'driver' in locals() and driver:
            driver.quit()


if __name__ == "__main__":
    main()
