import pandas as pd
import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager
from urllib.parse import urljoin

# --- Configuration ---

BASE_URL_FOR_JOINING = "https://www.lazarind.com"

CATEGORY_URLS = [
    "https://www.lazarind.com/collections/1",
    "https://www.lazarind.com/collections/2",
    "https://www.lazarind.com/collections/3",
    "https://www.lazarind.com/collections/4",
    "https://www.lazarind.com/collections/5",
    "https://www.lazarind.com/collections/6",
    "https://www.lazarind.com/collections/7",
    "https://www.lazarind.com/collections/8",
    "https://www.lazarind.com/collections/9",
    "https://www.lazarind.com/collections/10",
    "https://www.lazarind.com/collections/11",
    "https://www.lazarind.com/collections/12",
    "https://www.lazarind.com/collections/13",
    "https://www.lazarind.com/collections/14",
    "https://www.lazarind.com/collections/15",
    "https://www.lazarind.com/collections/16",
    "https://www.lazarind.com/collections/17", 
    "https://www.lazarind.com/collections/18",
    "https://www.lazarind.com/collections/19",
    "https://www.lazarind.com/collections/20",
    "https://www.lazarind.com/collections/21",
    "https://www.lazarind.com/collections/22",
    "https://www.lazarind.com/collections/23",
    "https://www.lazarind.com/collections/24",
    "https://www.lazarind.com/collections/30",
    "https://www.lazarind.com/collections/32",
    "https://www.lazarind.com/collections/34"
]

OUTPUT_FILE = "lazarind_product_DETAILS_Fixed.xlsx"
AUTOSAVE_LIMIT = 50 

def setup_driver():
    """Sets up a standard desktop Selenium WebDriver."""
    print("Setting up WebDriver...")
    chrome_options = webdriver.ChromeOptions()
    # chrome_options.add_argument("--headless") 
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--log-level=3")
    chrome_options.add_argument("start-maximized")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36")
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.maximize_window()
    print("WebDriver setup complete.")
    return driver

def save_data(data_list, filename):
    """Saves the collected data list to an Excel file."""
    if not data_list:
        print("No data to save.")
        return

    print(f"\nSaving {len(data_list)} products to {filename}...")
    try:
        df = pd.DataFrame(data_list)
        
        # Define the order of columns (Added 'Code' field)
        columns = [
            'Category', 'Product URL', 'Product Name', 'SKU', 'Code',
            'Dimension', 'Designer', 'Description', 'Overview', 
            'Grade', 'Color', 'Type', 'Content', 'Cleaning Code', 'Flame Code', 'Abrasion Rating',
            'Full Description (HTML)', 'Image URL'
        ]
        
        # Reorder DataFrame columns, adding any that might be missing
        df = df.reindex(columns=columns)
        
        df.to_excel(filename, index=False, engine='openpyxl')
        print(f"Successfully saved data to {filename}.")
    except Exception as e:
        print(f"CRITICAL: Failed to save data to {filename}. Error: {e}")
        try:
            fallback_file = filename.replace(".xlsx", ".csv")
            df.to_csv(fallback_file, index=False)
            print(f"WARNING: Saved data as a CSV fallback: {fallback_file}")
        except Exception as csv_e:
            print(f"CRITICAL: CSV fallback also failed. Error: {csv_e}")

def get_product_urls_from_category(driver, category_url):
    """
    Goes to a category page and scrapes the category name
    and a list of all product URLs on that page.
    """
    product_urls = []
    category_name = "Unknown Category"
    
    try:
        driver.get(category_url)
        
        # --- Get the category name ---
        try:
            category_name_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "h2.collection-title"))
            )
            category_name = category_name_element.text.strip()
        except Exception:
            try:
                category_name = driver.find_element(By.TAG_NAME, "h1").text.strip()
            except Exception:
                category_name = category_url.split('/')[-1]

        print(f"\n--- Scraping Category: '{category_name}' ({category_url}) ---")

        # --- Wait for items ---
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.k-listview-item"))
            )
        except TimeoutException:
            print(f"   Warning: Timed out waiting for products on {category_url}. Page may be empty.")
            return category_name, []

        # --- STALE ELEMENT FIX ---
        try:
            link_elements = driver.find_elements(By.CSS_SELECTOR, "div.k-listview-item a")
            
            raw_links = []
            for link in link_elements:
                try:
                    href = link.get_attribute('href')
                    if href:
                        raw_links.append(href)
                except StaleElementReferenceException:
                    continue 
            
            product_urls_seen = set()
            for relative_url in raw_links:
                product_url = urljoin(BASE_URL_FOR_JOINING, relative_url)
                if product_url not in product_urls_seen:
                    product_urls_seen.add(product_url)
                    product_urls.append(product_url)
                    
        except Exception as e:
             print(f"   Error extracting links: {e}")

        print(f"Found {len(product_urls)} product links for this category.")
        return category_name, product_urls

    except Exception as e:
        print(f"   Error scraping category {category_url}: {e}")
        return category_name, []

def get_product_details(driver, product_url, category_name):
    """
    Router function: Checks the URL structure and calls the appropriate scraper.
    """
    # Format 2: Fabric
    if "/fabrics/" in product_url:
        return _scrape_fabric_details(driver, product_url, category_name)
    # Format 3: Finishes (Added this check)
    elif "/finishes/" in product_url:
        return _scrape_finish_details(driver, product_url, category_name)
    # Format 1: Standard Furniture
    else:
        return _scrape_furniture_details(driver, product_url, category_name)

def _scrape_furniture_details(driver, product_url, category_name):
    """Original logic for standard furniture products (Format 1)."""
    driver.get(product_url)
    details = {
        'Category': category_name,
        'Product URL': product_url,
        'Product Name': "N/A",
        'SKU': "N/A",
        'Code': "N/A",
        'Dimension': "N/A",
        'Designer': "N/A",
        'Description': "N/A",
        'Overview': "N/A",
        'Full Description (HTML)': "N/A",
        'Image URL': "N/A"
    }
    
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "h2.mb-3"))
        )

        def safe_find_text(xpath, replace_str=None):
            try:
                elem = driver.find_element(By.XPATH, xpath)
                text = elem.text.strip()
                return text.replace(replace_str, "").strip() if replace_str else text
            except NoSuchElementException:
                return None

        details['Product Name'] = safe_find_text("//h2[contains(@class, 'mb-3')]")
        details['SKU'] = safe_find_text("//p[strong[text()='SKU:']]", "SKU:")
        details['Dimension'] = safe_find_text("//p[strong[text()='Dimension:']]", "Dimension:")
        details['Description'] = safe_find_text("//p[strong[text()='Description:']]", "Description:")
        details['Designer'] = safe_find_text("//a[strong[text()='Designer:']]", "Designer:")
        details['Overview'] = safe_find_text("//div[contains(@class, 'overview') and strong[text()='Overview:']]", "Overview:")

        try:
            full_desc_element = driver.find_element(By.XPATH, "//h2[contains(@class, 'mb-3')]/parent::div")
            details['Full Description (HTML)'] = full_desc_element.get_attribute('innerHTML').strip()
        except NoSuchElementException:
            details['Full Description (HTML)'] = None

        try:
            img_element = driver.find_element(By.CSS_SELECTOR, "div.f-carousel__slide[data-fancybox='gallery']")
            img_src = img_element.get_attribute('data-src') 
            if not img_src:
                img_tag = img_element.find_element(By.TAG_NAME, "img")
                img_src = img_tag.get_attribute('src')
            if img_src:
                details['Image URL'] = urljoin(BASE_URL_FOR_JOINING, img_src)
        except NoSuchElementException:
            details['Image URL'] = None
            
        return details

    except TimeoutException:
        print(f"   TIMEOUT on product page: {product_url}")
        return None
    except Exception as e:
        print(f"   ERROR scraping furniture {product_url}: {e}")
        return None

def _scrape_fabric_details(driver, product_url, category_name):
    """Logic for Fabric products (Format 2)."""
    driver.get(product_url)
    details = {
        'Category': category_name,
        'Product URL': product_url,
        'Product Name': "N/A",
        'SKU': "N/A",
        'Code': "N/A",
        'Grade': "N/A",
        'Color': "N/A",
        'Type': "N/A",
        'Content': "N/A",
        'Cleaning Code': "N/A",
        'Flame Code': "N/A",
        'Abrasion Rating': "N/A",
        'Full Description (HTML)': "N/A",
        'Image URL': "N/A"
    }

    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "h2.mb-2"))
        )

        try:
            details['Product Name'] = driver.find_element(By.CSS_SELECTOR, "h2.mb-2").text.strip()
        except:
            pass

        try:
            details['SKU'] = product_url.rstrip('/').split('/')[-1]
        except:
            pass

        def get_fabric_field(label):
            try:
                xpath = f"//*[contains(text(), '{label}')]"
                elem = driver.find_element(By.XPATH, xpath)
                full_text = elem.text.strip()
                return full_text.replace(label, "").strip()
            except NoSuchElementException:
                return "N/A"

        details['Grade'] = get_fabric_field("Grade:")
        details['Color'] = get_fabric_field("Color:")
        details['Type'] = get_fabric_field("Type:")
        details['Content'] = get_fabric_field("Content:")
        details['Cleaning Code'] = get_fabric_field("Cleaning Code:")
        details['Flame Code'] = get_fabric_field("Flame Code:")
        details['Abrasion Rating'] = get_fabric_field("Abrasion Rating:")

        try:
            container = driver.find_element(By.XPATH, "//h2[contains(@class,'mb-2')]/parent::div")
            details['Full Description (HTML)'] = container.get_attribute('innerHTML').strip()
        except NoSuchElementException:
            pass

        try:
            img_elem = driver.find_element(By.ID, "mainImage")
            src = img_elem.get_attribute('src')
            details['Image URL'] = urljoin(BASE_URL_FOR_JOINING, src)
        except NoSuchElementException:
            try:
                img_elem = driver.find_element(By.CSS_SELECTOR, "img.product-image")
                src = img_elem.get_attribute('src')
                details['Image URL'] = urljoin(BASE_URL_FOR_JOINING, src)
            except NoSuchElementException:
                pass

        return details

    except TimeoutException:
        print(f"   TIMEOUT on fabric page: {product_url}")
        return None
    except Exception as e:
        print(f"   ERROR scraping fabric {product_url}: {e}")
        return None

def _scrape_finish_details(driver, product_url, category_name):
    """New Logic for Finish products (Format 3)."""
    driver.get(product_url)
    details = {
        'Category': category_name,
        'Product URL': product_url,
        'Product Name': "N/A",
        'Code': "N/A", # New Field
        'Full Description (HTML)': "N/A",
        'Image URL': "N/A"
    }

    try:
        # Wait for the Image ID to be present (as it seems consistent in your request)
        # or the H2
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "mainImage"))
        )

        # 1. Product Name (Assuming H2 based on usual patterns, or grabbing from URL/Title if H2 varies)
        try:
            # Trying generic H2 first as you mentioned "as usual"
            details['Product Name'] = driver.find_element(By.TAG_NAME, "h2").text.strip()
        except NoSuchElementException:
             pass

        # 2. Extract Code
        # Target: <p class="text-muted mb-4">Code ODN</p>
        try:
            code_element = driver.find_element(By.XPATH, "//p[contains(@class, 'text-muted') and contains(text(), 'Code')]")
            code_text = code_element.text.strip()
            # Clean "Code " from the string to get just "ODN"
            details['Code'] = code_text.replace("Code", "").strip()
        except NoSuchElementException:
            details['Code'] = "N/A"

        # 3. Full Description (HTML) - Assuming parent of H2/Code area
        try:
            # Often the description is the parent of the code element or the H2
            container = driver.find_element(By.XPATH, "//p[contains(@class, 'text-muted') and contains(text(), 'Code')]/parent::div")
            details['Full Description (HTML)'] = container.get_attribute('innerHTML').strip()
        except NoSuchElementException:
            pass

        # 4. Image URL
        # Target: <img ... id="mainImage">
        try:
            img_elem = driver.find_element(By.ID, "mainImage")
            src = img_elem.get_attribute('src')
            details['Image URL'] = urljoin(BASE_URL_FOR_JOINING, src)
        except NoSuchElementException:
            pass

        return details

    except TimeoutException:
        print(f"   TIMEOUT on finish page: {product_url}")
        return None
    except Exception as e:
        print(f"   ERROR scraping finish {product_url}: {e}")
        return None

def main():
    driver = setup_driver()
    all_products_data = []
    total_products_scraped_session = 0
    
    try:
        print(f"\n--- Starting Scrape of {len(CATEGORY_URLS)} Categories ---")
        
        for category_url in CATEGORY_URLS:
            category_name, product_urls = get_product_urls_from_category(driver, category_url)
            total_in_category = len(product_urls)
            
            if not product_urls:
                continue
                
            for i, product_url in enumerate(product_urls):
                
                # Checks URL type and calls specific scraper
                details = get_product_details(driver, product_url, category_name)
                
                if details:
                    all_products_data.append(details)
                    total_products_scraped_session += 1

                    name = details.get('Product Name', 'N/A')
                    print(f"   Scraping [{i+1} / {total_in_category}] -> {name} -> {product_url}")

                    if total_products_scraped_session % AUTOSAVE_LIMIT == 0 and total_products_scraped_session > 0:
                        print(f"\n--- Auto-saving data ({total_products_scraped_session} total products scraped) ---")
                        save_data(all_products_data, OUTPUT_FILE)
                        print("--- Resuming scrape ---\n")
                
                time.sleep(0.2) 

    except KeyboardInterrupt:
        print("\n--- User interrupted (Ctrl+C). Saving data before exiting. ---")
    except Exception as e:
        print(f"\n--- An unexpected error occurred: {e} ---")
        print("--- Saving all collected data before exiting. ---")
    
    finally:
        if driver:
            driver.quit()
            print("WebDriver closed.")
        
        if all_products_data:
            print(f"\n--- Final Save ---")
            save_data(all_products_data, OUTPUT_FILE)
        else:
            print("\nNo new data was collected in this session.")

if __name__ == "__main__":
    main()