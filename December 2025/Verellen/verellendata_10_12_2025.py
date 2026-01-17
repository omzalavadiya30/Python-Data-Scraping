import time
import pandas as pd
import re
import signal
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# Global flag for interrupt handling
stop_script = False

def signal_handler(sig, frame):
    global stop_script
    print("\n\n⚠️ Interrupt received (Ctrl+C). Stopping safely after current item...")
    stop_script = True

signal.signal(signal.SIGINT, signal_handler)

class VerellenScraper:
    def __init__(self):
        options = Options()
        options.add_argument("--start-maximized")
        # options.add_argument("--headless") 
        self.driver = webdriver.Chrome(options=options)
        self.wait = WebDriverWait(self.driver, 20) # Increased wait time to 20s
        self.master_data = []
        self.base_url = "https://verellen.biz/"
        self.items_scraped_total = 0

    # --- MENU & NAVIGATION LOGIC ---
    def reset_to_main_menu(self):
        try:
            menu_icon = self.driver.find_element(By.CSS_SELECTOR, "div.icon-holder.menu")
            if "closed" in menu_icon.get_attribute("class"):
                menu_icon.click()
                self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.icon-holder.menu.open")))
                time.sleep(1)
        except: pass

    def navigate_to_submenu_2(self, main_category_text):
        self.reset_to_main_menu()
        try:
            xpath = f"//div[@class='items-wrapper wrapper-with-scroller']//a[contains(text(), '{main_category_text}')]"
            link = self.driver.find_element(By.XPATH, xpath)
            link.click()
            self.wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.submenu-2.appear")))
            time.sleep(1)
            return True
        except: return False

    def get_category_structure(self):
        category_links = []
        
        # 1. Attic
        self.driver.get(self.base_url)
        time.sleep(2)
        self.reset_to_main_menu()
        try:
            attic_link = self.driver.find_element(By.XPATH, "//div[@class='items-wrapper wrapper-with-scroller']//a[contains(text(), 'Attic')]")
            category_links.append({"Category": "Attic", "SubCategory": "Main", "URL": attic_link.get_attribute("href")})
        except: pass

        # 2. Products & Materials
        for main_cat in ["Products", "Materials and Finishes"]:
            if not self.navigate_to_submenu_2(main_cat): continue
            try:
                submenu2 = self.driver.find_element(By.CSS_SELECTOR, "div.submenu-2.appear .items-wrapper")
                item_names = [x.text for x in submenu2.find_elements(By.TAG_NAME, "a") if x.text and "View All" not in x.text]
            except: continue

            for item_name in item_names:
                self.navigate_to_submenu_2(main_cat)
                try:
                    clean_name = item_name.split('&')[0].strip()
                    xpath = f"//div[contains(@class, 'submenu-2')]//a[contains(text(), \"{clean_name}\")]"
                    element = self.driver.find_element(By.XPATH, xpath)
                    direct_href = element.get_attribute("href")
                    element.click()
                    time.sleep(1.5)

                    try:
                        submenu3 = self.driver.find_element(By.CSS_SELECTOR, "div.submenu-3.appear")
                        if submenu3.is_displayed():
                            for sub3 in submenu3.find_elements(By.CSS_SELECTOR, ".items-wrapper a"):
                                category_links.append({
                                    "Category": main_cat,
                                    "SubCategory": f"{item_name} - {sub3.text}",
                                    "URL": sub3.get_attribute("href")
                                })
                            continue
                    except: pass 
                    category_links.append({"Category": main_cat, "SubCategory": item_name, "URL": direct_href})
                except: pass
        return category_links

    # --- LISTING PAGE SCRAPER ---
    def scrape_listing_urls(self, url):
        product_urls = set()
        try:
            self.driver.get(url)
            time.sleep(3)
            
            while True:
                container = self.driver.find_element(By.CSS_SELECTOR, "div.plp-listing-container")
                children = container.find_elements(By.XPATH, "./div")
                
                new_found = False
                for child in children:
                    try:
                        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", child)
                        time.sleep(0.1) 
                    except: pass

                    p_url = None
                    try:
                        link_el = child.find_element(By.CSS_SELECTOR, "a.product-name")
                        p_url = link_el.get_attribute("href")
                    except: 
                        time.sleep(0.3)
                        try:
                            link_el = child.find_element(By.CSS_SELECTOR, "a.product-name")
                            p_url = link_el.get_attribute("href")
                        except: pass
                    
                    if p_url and p_url not in product_urls:
                        product_urls.add(p_url)
                        new_found = True

                try:
                    next_btn = self.driver.find_element(By.XPATH, "//div[contains(@class, 'prev-page')]//p[contains(text(), 'Next')]")
                    if next_btn.is_displayed():
                        self.driver.execute_script("arguments[0].click();", next_btn)
                        time.sleep(4)
                    else: break
                except: break
                
        except Exception as e:
            print(f"    Error collecting URLs: {e}")
        
        return list(product_urls)

    # --- PRODUCT DETAILS HELPER ---
    def extract_text(self, xpath, default="N/A"):
        try:
            return self.driver.find_element(By.XPATH, xpath).text.strip()
        except: return default

    def extract_breadcrumbs(self):
        try:
            crumbs = self.driver.find_elements(By.CSS_SELECTOR, "div.breadcrumbs ul li a span")
            return " | ".join([c.text.strip() for c in crumbs])
        except: return "N/A"

    def extract_sku(self, category_type):
        try:
            crumbs = self.driver.find_elements(By.CSS_SELECTOR, "div.breadcrumbs ul li a span")
            if not crumbs: return "N/A"
            last_crumb = crumbs[-1].text.strip()

            if category_type == "Attic":
                if any(char.isdigit() for char in last_crumb):
                    return last_crumb
                return "N/A"
            elif category_type == "Products":
                return last_crumb.replace("(", "").replace(")", "").strip()
            else: return "N/A" 
        except: return "N/A"

    def extract_specs(self, target_map):
        try:
            # Look for rows in both Furniture and Material styles
            rows = self.driver.find_elements(By.CSS_SELECTOR, "div.content-detail, div.d-grid")
            found_data = {}
            for row in rows:
                try:
                    cols = row.find_elements(By.XPATH, "./div | ./p")
                    # Materials often use <p>Label</p><p>Value</p> inside a d-grid
                    # Furniture often uses <div><p>Label</p></div><div><p>Value</p></div> inside content-detail
                    
                    label = ""
                    value = ""

                    if len(cols) >= 2:
                        # Attempt to get text from first col (Label)
                        label = cols[0].text.replace(":", "").strip()
                        
                        # Attempt to get text from second col (Value)
                        val_el = cols[1]
                        # Check if value element has nested <p> tags (common in dimensions)
                        nested_ps = val_el.find_elements(By.TAG_NAME, "p")
                        if nested_ps:
                            value = " | ".join([p.text.strip() for p in nested_ps if p.text.strip()])
                        else:
                            value = val_el.text.strip()

                        if label and value:
                            found_data[label.lower()] = value
                except: continue

            # Map to target keys
            result = {}
            for key in target_map:
                # Use lower case for matching to be safe
                result[key] = found_data.get(key.lower(), "N/A")
            return result
        except:
            return {k: "N/A" for k in target_map}

    def scrape_product_details(self, url, category_context, main_cat_type):
        attempts = 0
        while attempts < 2: # Retry logic
            try:
                self.driver.get(url)
                
                # --- CRITICAL WAIT ---
                # Wait up to 20 seconds for the Product Name to be visible
                self.wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.product-name h1")))
                
                # Wait extra time for JS to load Breadcrumbs and Slick Slider
                time.sleep(2) 

                product_data = {
                    "Category Hierarchy": self.extract_breadcrumbs(),
                    "Product Name": self.extract_text("//div[contains(@class,'product-name')]/h1"),
                    "SKU": self.extract_sku(main_cat_type),
                    "Product URL": url,
                    "Price": "N/A",
                    "Brand": "Verellen",
                    "Description": self.extract_text("//p[contains(@class, 'mt-3')]")
                }

                try:
                    price_el = self.driver.find_element(By.XPATH, "//span[@itemprop='price']")
                    product_data["Price"] = price_el.text.strip()
                except: pass

                # --- SPECS ---
                if "Materials" in category_context or "Materials" in main_cat_type:
                    mat_fields = [
                        "Grade", "Color", "Leather type", "Hide size", "Thickness", "Content", 
                        "Repeat", "Width", "Application", "Abrasion", "Testing", "Country of Origin", 
                        "Available Stitch Detail", "Cleaning Code", "Available On Lola Lounge", 
                        "Special Details", "Contract", "Heavy Duty Abrasion Rating", "OEKOTEX Certification"
                    ]
                    specs = self.extract_specs(mat_fields)
                    product_data.update(specs)
                else:
                    dim_fields = [
                        "Overall Width", "Overall Height", "Overall Length", "Overall Depth", 
                        "Overall Diameter", "Base Height", "Top Thickness", "Interior Depth", 
                        "Seat Depth", "Seat Height", "Arm Height", "Arm Width", "Leg Height"
                    ]
                    specs = self.extract_specs(dim_fields)
                    product_data.update(specs)

                # --- IMAGES ---
                try:
                    # Target slick-slide inside slick-track that are not cloned
                    images = self.driver.find_elements(By.CSS_SELECTOR, ".slick-track .slick-slide:not(.slick-cloned) img")
                    img_urls = []
                    for img in images:
                        src = img.get_attribute("src")
                        if src and "placeholder" not in src: # Filter placeholders if any
                             img_urls.append(src)
                    
                    # Remove duplicates
                    seen = set()
                    unique_imgs = [x for x in img_urls if not (x in seen or seen.add(x))]

                    for i in range(4):
                        col_name = f"Image{i+1}"
                        product_data[col_name] = unique_imgs[i] if i < len(unique_imgs) else "N/A"
                except:
                    for i in range(4): product_data[f"Image{i+1}"] = "N/A"

                return product_data

            except TimeoutException:
                print(f"    ⚠️ Timeout loading {url}. Retrying...")
                attempts += 1
                time.sleep(2)
            except Exception as e:
                print(f"    ❌ Error scraping details: {e}")
                return None
        
        return None # Failed after retries

    # --- MAIN ---
    def save_data(self):
        if not self.master_data: return
        df = pd.DataFrame(self.master_data)
        filename = "verellen_full_details.xlsx"
        df.to_excel(filename, index=False)
        print(f"\n💾 Saved {len(self.master_data)} records to {filename}")

    def run(self):
        global stop_script
        try:
            print("--- Step 1: Mapping Categories ---")
            structure = self.get_category_structure()
            print(f"✅ Found {len(structure)} sub-categories.")
            
            for entry in structure:
                if stop_script: break
                
                cat_name = entry['SubCategory']
                main_cat = entry['Category']
                listing_url = entry['URL']
                
                print(f"\n📂 Processing Category: {cat_name} ({main_cat})")
                
                # 1. Get URLs
                product_urls = self.scrape_listing_urls(listing_url)
                total_in_cat = len(product_urls)
                print(f"   > Found {total_in_cat} products.")

                # 2. Scrape Details
                for idx, p_url in enumerate(product_urls):
                    if stop_script: break
                    
                    cat_type_logic = "Materials" if "Materials" in main_cat else ("Attic" if "Attic" in main_cat else "Products")
                    
                    details = self.scrape_product_details(p_url, main_cat, cat_type_logic)
                    
                    if details:
                        details['Source Category'] = cat_name
                        self.master_data.append(details)
                        self.items_scraped_total += 1
                        
                        # Logging format requested
                        p_name = details.get('Product Name', 'Unknown')
                        p_sku = details.get('SKU', 'N/A')
                        print(f"   Scraping {idx+1}/{total_in_cat} -> {p_name} -> {p_sku} -> {p_url}")

                    # Incremental Save
                    if self.items_scraped_total % 50 == 0:
                        self.save_data()

            self.save_data()

        except KeyboardInterrupt:
            print("\n🛑 Manual Interruption.")
        finally:
            self.save_data()
            self.driver.quit()

if __name__ == "__main__":
    bot = VerellenScraper()
    bot.run()