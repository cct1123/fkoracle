"""
Oracle Procurement Automation Pipeline

This script parses a Thorlabs shopping cart CSV export and injects the parsed 
inventory data directly into the Oracle Application Development Framework (ADF) 
Non-Catalog Request form via Selenium WebDriver.
"""

import csv
import json
import time

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException


# ==========================================
# Data Ingestion & Parsing
# ==========================================

def parse_thorlabs_cart_to_target(file_path: str) -> dict:
    """
    Parses a Thorlabs CSV cart export into a structured dictionary optimized 
    for the Oracle Procurement pipeline.

    Args:
        file_path (str): The local system path to the Thorlabs CSV file.

    Returns:
        dict: A dictionary containing the parsed inventory items with their 
              descriptions, quantities, prices, and line amounts.
    """
    quote_dictionary = {"inventory": {}}
    
    # utf-8-sig automatically handles the Byte Order Mark (BOM) common in Excel exports
    with open(file_path, mode='r', encoding='utf-8-sig') as file:
        reader = csv.DictReader(file)
        
        for row in reader:
            part_number = row.get("Item Number", "").strip()
            
            # Skip empty rows or the standard Thorlabs CSV disclaimer
            if not part_number or "This CSV export" in part_number:
                continue
                
            try:
                description = row.get("Description", "").strip()
                quantity = float(row.get("Quantity", 0))
                
                # Clean currency strings (e.g., "$100.91" -> 100.91)
                raw_price = row.get("Unit Price", "$0.0").replace('$', '').replace(',', '')
                price = float(raw_price)
                
                raw_amount = row.get("Line Total", "$0.0").replace('$', '').replace(',', '')
                amount = float(raw_amount)
                
                # Map to the target schema
                quote_dictionary["inventory"][part_number] = {
                    "description": description,
                    "quantity": quantity,
                    "price": price,
                    "amount": amount
                }
                
            except ValueError as e:
                print(f"Data type conversion error on part {part_number}: {e}")
                continue
                
    return quote_dictionary


# ==========================================
# WebDriver Initialization
# ==========================================

def create_active_session(target_url: str):
    """
    Initializes a Chrome WebDriver session configured for enterprise 
    application automation.

    Args:
        target_url (str): The target URL to load upon browser launch.

    Returns:
        webdriver.Chrome | None: An active WebDriver instance, or None if initialization fails.
    """
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-infobars")
    chrome_options.add_argument("--disable-extensions")

    try:
        print("Initializing WebDriver session...")
        driver = webdriver.Chrome(options=chrome_options)
        
        # Note: Implicit waits are deliberately minimized to prevent 
        # conflicts with Explicit Waits (WebDriverWait) used later in the script.
        driver.implicitly_wait(10)

        print(f"Navigating to: {target_url}")
        driver.get(target_url)

        return driver

    except Exception as e:
        print(f"Failed to establish WebDriver session: {e}")
        return None


# ==========================================
# Oracle Automation Controller
# ==========================================

class OracleProcurementAutomator:
    """
    A Selenium-based automation controller designed to inject structured quote 
    dictionary data into the Oracle ADF 'Non-Catalog Request' form.

    Attributes:
        driver (webdriver.Chrome): The active browser session.
        wait (WebDriverWait): The explicit wait configuration.
        cached_cart_btn (WebElement): In-memory cache for the Add to Cart button.
        locators (dict): XPath strategies anchored on visible labels.
    """

    def __init__(self, driver, timeout=15):
        """
        Initializes the automator with wait strategies and DOM locators.

        Args:
            driver: The active Selenium WebDriver instance.
            timeout (int): Maximum seconds to wait for element conditions.
        """
        self.driver = driver
        self.wait = WebDriverWait(
            self.driver, 
            timeout, 
            poll_frequency=0.5,
            ignored_exceptions=[StaleElementReferenceException]
        )
        self.cached_cart_btn = None
        
        # XPath strategies anchoring on visible labels to bypass dynamic ADF IDs
        self.locators = {
            "item_description": (By.XPATH, "//label[contains(text(), 'Item Description')]/ancestor::tr[1]//textarea"),
            "category_name": (By.XPATH, "//label[contains(text(), 'Category Name')]/ancestor::tr[1]//input[not(@type='hidden')]"),
            "quantity": (By.XPATH, "//label[contains(text(), 'Quantity')]/ancestor::tr[1]//input[not(@type='hidden')]"),
            "uom_name": (By.XPATH, "//label[contains(text(), 'UOM Name')]/ancestor::tr[1]//input[not(@type='hidden')]"),
            "price": (By.XPATH, "//label[contains(text(), 'Price')]/ancestor::tr[1]//input[not(@type='hidden')]"),
            "supplier": (By.XPATH, "//label[contains(text(), 'Supplier')]/ancestor::tr[1]//input[not(@type='hidden')]"),
            "supplier_site": (By.XPATH, "//label[contains(text(), 'Supplier Site')]/ancestor::tr[1]//input[not(@type='hidden')]"),
            "add_to_cart_btn": (By.XPATH, "//button[contains(text(), 'Add to Cart')]")
        }

    def _input_text(self, locator_key: str, text: str):
        """
        Waits for element interactability, clears existing data, and injects new text.

        Args:
            locator_key (str): The key mapping to the self.locators dictionary.
            text (str): The string data to inject.
        """
        try:
            element = self.wait.until(EC.element_to_be_clickable(self.locators[locator_key]))
            element.clear()
            element.send_keys(str(text))
        except TimeoutException:
            print(f"Timeout: Could not locate or interact with {locator_key}.")

    def _click_empty_space(self):
        """
        Executes a geometric click at the absolute (5,5) pixel offset of the document body.
        Forces an input 'blur' event, instantly collapsing any active ADF auto-suggest 
        dropdowns or List of Values (LOV) overlays.
        """
        try:
            body_element = self.driver.find_element(By.TAG_NAME, "body")
            action = ActionChains(self.driver)
            action.move_to_element_with_offset(body_element, 5, 5).click().perform()
        except Exception as e:
            print(f"Warning: Neutral space click failed. Error: {e}")

    def process_quote_inventory(self, quote_dictionary: dict, supplier_name: str, supplier_site: str, category: str):
        """
        Iterates through the extracted hardware dictionary and populates the 
        Oracle Procurement UI sequentially.

        Args:
            quote_dictionary (dict): The parsed inventory data.
            supplier_name (str): The target supplier name for the ADF form.
            supplier_site (str): The specific site location for the supplier.
            category (str): The overarching procurement category classification.
        """
        inventory = quote_dictionary.get("inventory", {})

        for part_number, details in inventory.items():
            print(f"Injecting Part: {part_number}...")

            # Description (Note: switched to single quotes inside f-string interpolation)
            self._input_text("item_description", f"Part number: {part_number} \n {details['description']}")
            self._click_empty_space()

            # Category
            self._input_text("category_name", category)
            self._click_empty_space()

            # Quantity
            self._input_text("quantity", details["quantity"])
            self._click_empty_space()

            # Price
            price = details["price"]
            self._input_text("price", price)
            self._click_empty_space()

            # Supplier (Triggers primary background validation)
            self._input_text("supplier", supplier_name)
            self._click_empty_space()

            # Supplier Site (Commented out in provided prompt, preserved as-is)
            # self._input_text("supplier_site", supplier_site)
            # self._click_empty_space()

            # Commit Item
            self._fast_click_add_to_cart()
            self._wait_for_adf_stabilization()

    def _fast_click_add_to_cart(self):
        """
        Attempts to click using a cached memory reference to bypass DOM searching.
        Recovers automatically if the ADF framework re-rendered the button node.
        """
        cart_xpath = (
            "//button[contains(normalize-space(), 'Add to Cart')] | "
            "//a[.//span[contains(normalize-space(), 'Add to Cart')]] | "
            "//a[contains(normalize-space(), 'Add to Cart')]"
        )

        # 1. Attempt to use the cached element (Fastest path)
        if self.cached_cart_btn:
            try:
                self.driver.execute_script("arguments[0].click();", self.cached_cart_btn)
                return 
            except StaleElementReferenceException:
                print("Cache stale: ADF re-rendered the button. Re-locating...")
                self.cached_cart_btn = None

        # 2. Locate, Cache, and Click (Executed on 1st item, or if cache goes stale)
        try:
            self.cached_cart_btn = self.wait.until(EC.presence_of_element_located((By.XPATH, cart_xpath)))
            self.driver.execute_script("arguments[0].click();", self.cached_cart_btn)
        except TimeoutException:
            print("Timeout: 'Add to Cart' button not found.")

    def _wait_for_adf_stabilization(self):
        """
        Pauses execution until the Oracle ADF loader overlay disappears.
        Monitors the 'Item Description' field until it clears, indicating 
        the form has successfully reset for the next entry.
        """
        try:
            time.sleep(1) # Brief pause to allow ADF to trigger form reset
            self.wait.until(
                lambda d: d.find_element(*self.locators["item_description"]).get_attribute("value") == ""
            )
        except TimeoutException:
            print("Warning: Form did not reset within timeout parameters.")


# ==========================================
# Execution Pipeline
# ==========================================

if __name__ == "__main__":
    target_csv = "2026-03-27-Thorlabs-Cart.csv"
    oracle_url = "https://fa-eusf-saasfaprod1.fa.ocs.oraclecloud.com"
    
    # 1. Parse the local spreadsheet data
    try:
        parsed_dictionary = parse_thorlabs_cart_to_target(target_csv)
        print("CSV Parsed Successfully. Sample structure:")
        print(json.dumps(parsed_dictionary, indent=4))
    except FileNotFoundError:
        print(f"System Error: File '{target_csv}' not found. Exiting.")
        exit(1)

    # 2. Establish browser connection
    active_driver = create_active_session(oracle_url)

    if active_driver:
        try:
            print("Session active. Ready for automation logic.")
            print(f"Current Title: {active_driver.title}")

            # 3. Allow human operator to navigate SSO and initial menus
            print("Please navigate to the classic non-catalog page.")
            input("Press Enter to initiate data injection...")

            # 4. Execute the automation payload
            automator = OracleProcurementAutomator(active_driver)
            automator.process_quote_inventory(
                quote_dictionary=parsed_dictionary, 
                supplier_name="Thorlabs Inc", 
                supplier_site="43 SPARTA AVE", 
                category="LABORATORY SUPPLIES <$5k"
            )
            
            print("Procurement pipeline execution complete.")

        finally:
            # 5. Clean teardown
            print("Closing WebDriver session.")
            active_driver.quit()
