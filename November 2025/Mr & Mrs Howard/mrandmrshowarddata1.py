import pandas as pd
import time
import sys
import signal
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# --- GLOBAL VARIABLES FOR DATA SAFETY ---
all_data = []
output_file = "sherrill_furniture_complete_data.xlsx"

# --- SIGNAL HANDLER FOR CTRL+C ---
def signal_handler(sig, frame):
    print("\n\n[!] Interruption detected (Ctrl+C). Saving collected data...")
    save_data(all_data)
    sys.exit(0)

signal.signal(signal.SIGINT, signal_handler)

# --- SAVE FUNCTION ---
def save_data(data):
    if not data:
        print("No data to save.")
        return
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False)
    print(f"[✓] Data saved to '{output_file}' ({len(data)} records)")

def scrape_sherrill_furniture():
    # Setup Chrome Driver
    options = webdriver.ChromeOptions()
    options.add_argument('--start-maximized')
    # options.add_argument('--headless') # Uncomment to run in background
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 15)
    
    start_url = "https://mrandmrshoward.sherrillfurniture.com/search_results.php"
    product_urls = []
    
    # ==========================================
    # PHASE 1: COLLECT PRODUCT URLS
    # ==========================================
    try:
        print(f"--- PHASE 1: Collecting URLs from {start_url} ---")
        driver.get(start_url)
        
        # Handle "View All"
        try:
            view_all_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#view-all a")))
            print("Found 'View All' button. Clicking...")
            view_all_btn.click()
            wait.until(EC.url_contains("view=all"))
            time.sleep(5)
        except Exception as e:
            print("Proceeding with current view (View All not found or already active).")

        # Extract Links
        elements = driver.find_elements(By.TAG_NAME, "a")
        for elem in elements:
            try:
                href = elem.get_attribute("href")
                if href and "search_results_detail.php" in href:
                    if href not in product_urls:
                        product_urls.append(href)
            except:
                continue

        print(f"Total unique products found: {len(product_urls)}")
        
    except Exception as e:
        print(f"Error in Phase 1: {e}")
        driver.quit()
        return

    # ==========================================
    # PHASE 2: EXTRACT PRODUCT DETAILS
    # ==========================================
    print(f"--- PHASE 2: Scraping Details for {len(product_urls)} Products ---")
    
    count = 0
    total_products = len(product_urls)

    for url in product_urls:
        count += 1
        
        # Initialize default values
        item_data = {
            "Category": "Products",
            "Product Url": url,
            "Product Name": "N/A",
            "SKU": "N/A",
            "Brand": "Mr. and Mrs. Howard for Sherrill Furniture",
            "Description": "N/A",
            "ItemPdf": "N/A",
            "Image1": "N/A",
            "Image2": "N/A",
            "Image3": "N/A",
            "Image4": "N/A"
        }

        try:
            driver.get(url)
            
            # Wait for a key element to ensure page load (Description or Footer)
            try:
                wait.until(EC.presence_of_element_located((By.ID, "result-description")))
            except:
                time.sleep(2) # Fallback wait

            # 1. SKU and Name Extraction (IMPROVED)
            try:
                # Get ALL h2 elements
                h2_list = driver.find_elements(By.TAG_NAME, "h2")
                header_text = ""
                
                # Loop to find the first H2 that actually has text
                for h2 in h2_list:
                    # 'textContent' gets text even if hidden by CSS
                    txt = h2.get_attribute("textContent").strip()
                    if txt: 
                        header_text = txt
                        break 
                
                if header_text:
                    # Split logic: "H110C Rhian Armless Chair"
                    if " " in header_text:
                        parts = header_text.split(" ", 1)
                        item_data["SKU"] = parts[0].strip()
                        item_data["Product Name"] = parts[1].strip()
                    else:
                        item_data["SKU"] = header_text
                        item_data["Product Name"] = "N/A"
                else:
                    print(" [!] Warning: No text found in any H2 tag.")
            except Exception as e:
                print(f" [!] Error extracting Name/SKU: {e}")

            # 2. Description (Full HTML)
            try:
                desc_elem = driver.find_element(By.ID, "result-description")
                item_data["Description"] = desc_elem.get_attribute("outerHTML").strip()
            except:
                pass

            # 3. Item PDF
            try:
                # Force scroll to bottom to trigger any lazy loading
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                
                # Look for the specific link inside the user-menu list
                pdf_links = driver.find_elements(By.CSS_SELECTOR, "ul.user-menu a[href*='search_results_pdf.php']")
                if pdf_links:
                    item_data["ItemPdf"] = pdf_links[0].get_attribute("href")
            except:
                pass

            # 4. Images
            # Image 1: Main popup image
            try:
                img1_elem = driver.find_elements(By.CSS_SELECTOR, "#search-result-photo a.item-popup-image")
                if img1_elem:
                    item_data["Image1"] = img1_elem[0].get_attribute("href")
            except:
                pass

            # Image 2-4: From #roomscene
            try:
                room_imgs = driver.find_elements(By.CSS_SELECTOR, "#roomscene a.item-popup-image")
                for i, img_elem in enumerate(room_imgs):
                    if i == 0: item_data["Image2"] = img_elem.get_attribute("href")
                    if i == 1: item_data["Image3"] = img_elem.get_attribute("href")
                    if i == 2: item_data["Image4"] = img_elem.get_attribute("href")
                    if i >= 2: break 
            except:
                pass

            # Add to global list
            all_data.append(item_data)

            # LOGGING
            print(f"Scraping {count}/{total_products} -> {item_data['Product Name']} -> {item_data['SKU']} -> {url}")

            # AUTO-SAVE every 50 records
            if count % 50 == 0:
                print(f"[i] Checkpoint reached. Saving data...")
                save_data(all_data)

        except Exception as e:
            print(f"Error scraping {url}: {e}")
            all_data.append(item_data)

    # Final Save
    driver.quit()
    print("--- Scraping Completed ---")
    save_data(all_data)

if __name__ == "__main__":
    scrape_sherrill_furniture()