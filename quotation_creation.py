import pandas as pd
from datetime import datetime
from selenium import webdriver
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import logging # Import the logging module
import random

# Create a FileHandler to write logs to a file
log_file_path = 'automation.log'
file_handler = logging.FileHandler(log_file_path, mode='w') # 'a' for append mode, w for write
file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

# Create a StreamHandler to print logs to the console
console_handler = logging.StreamHandler()
console_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

logging.basicConfig(level=logging.INFO, # Set overall logging level
                    handlers=[file_handler, console_handler]) # Add both handlers

logger = logging.getLogger(__name__)

# --- Configuration Constants ---
EXCEL_INPUT_FILE = "clienti_assicurazioni.xlsx"
EXCEL_OUTPUT_FILE = "customers_updated.xlsx"
START_DATE_STR = "2025-07-12"
END_DATE_STR = "2025-07-15"
PRIMA_LOGIN_URL = "https://intermediari.prima.it/login"
PRIMA_USER = "prima@pass-broker.it"
PRIMA_PASSWORD = "Prima2025!"

# --- Selenium Locators (Define them once for readability and maintainability) ---
# Login Page
LOC_USERNAME_INPUT = (By.ID, "id-input-email-id")
LOC_PASSWORD_INPUT = (By.ID, "id-input-password-id")
LOC_LOGIN_BUTTON = (By.CSS_SELECTOR, "button[type='submit']")

# Dashboard / Navigation
LOC_QUOTATION_BUTTON = (By.XPATH,"//a[@class='quotation-link']")
LOC_COMPUTE_MOTOR_BUTTON = (By.CSS_SELECTOR, "button[data-test-id='calcola-motor']")

# Quotation Form - Page 1 (Vehicle/Owner Basic Data)
LOC_PLATE_INPUT = (By.ID, "plate_number")
LOC_BIRTHDAY_INPUT = (By.XPATH, "(//input[@id='owner_birth_date'])[2]") # Be careful with index if multiple exist
LOC_PROCEED_BUTTON = (By.XPATH, "(//button[@type='button'])[1]") # This is very generic, consider a better one if possible

# Quotation Form - Page 2 (Additional Owner Data)
LOC_EFFECTIVE_DATE_INPUT = (By.XPATH, "(//input[@id='effective_date_date'])[2]") # Check this path again
LOC_LICENSE_YEAR_DROPDOWN = (By.CSS_SELECTOR, "div[id='owner_license_year'] span[class='form-select__status']")
LOC_LICENSE_YEAR_OPTION = lambda year: (By.XPATH, f"//div[@id='owner_license_year']//li[normalize-space()='{year}']")
LOC_DEFAULT_LICENSE_YEAR_OPTION = (By.CSS_SELECTOR, "div[id='owner_license_year'] li:nth-child(1)")

LOC_RESIDENTIAL_CITY_INPUT = (By.ID, "owner_residential_city")
LOC_RESIDENTIAL_CITY_SELECT_FIRST_OPTION = (By.CSS_SELECTOR, "div[class='is-valid form-autocomplete is-open is-large is-pristine'] li:nth-child(1)") # More general class

LOC_CAP_INPUT = (By.ID, "owner_residential_cap")
LOC_CAP_SELECT_OPTION = lambda cap: (By.XPATH, f"(//li[normalize-space()='{cap}'])[1]") # Assuming CAP has a select dropdown too
LOC_DEFAULT_CAP_OPTION = (By.XPATH, "(//li[normalize-space()='10121'])[1]") # Default to Turin CAP

LOC_RESIDENTIAL_ADDRESS_INPUT = (By.ID, "owner_residential_address")
LOC_RESIDENTIAL_NUMBER_INPUT = (By.ID, "owner_residential_civic_number")

LOC_OCCUPATION_DROPDOWN = (By.XPATH, "(//div[@id='owner_occupation'])[1]")
LOC_OCCUPATION_SELECT_SECOND_OPTION = (By.CSS_SELECTOR, "div[id='owner_occupation'] li:nth-child(2)")

LOC_CIVIL_STATUS_DROPDOWN = (By.XPATH, "(//div[@id='owner_civil_status'])[1]")
LOC_CIVIL_STATUS_SELECT_FIRST_OPTION = (By.CSS_SELECTOR, "div[id='owner_civil_status'] li:nth-child(1)")

LOC_CELL_NUMBER_INPUT = (By.CSS_SELECTOR, "#phone_number")
LOC_PRIVACY_CHECKBOX = (By.CSS_SELECTOR, "label[for='privacy_all']")
LOC_COMPUTE_QUOTATION_BUTTON = (By.CSS_SELECTOR, ".btn.btn--primary[data-test-id='button-calculate-quote']")

# Quotation Result Page
LOC_QUOTATION_PRICE = (By.CSS_SELECTOR, "div[class='guarantee-box__price guarantee-box__price--highlighted'] span[class='price__value']")


# --- Helper Functions for Selenium Interactions ---

def wait_and_click(driver, locator, timeout=5):
    """Waits for an element to be clickable and then clicks it."""
    try:
        element = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable(locator)
        )
        element.click()
        logger.debug(f"Clicked element: {locator}")
        return True
    except Exception as e:
        logger.warning(f"Could not click {locator} within {timeout}s: {e}")
        return False

def wait_and_send_keys(driver, locator, keys, timeout=5):
    """Waits for an element to be present and sends keys to it."""
    try:
        element = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located(locator)
        )
        element.send_keys(keys)
        logger.debug(f"Sent keys to {locator}: '{keys}'")
        return True
    except Exception as e:
        logger.warning(f"Could not send keys to {locator} within {timeout}s: {e}")
        return False

def scroll_to_element(driver, locator, timeout=5):
    """Scrolls to an element."""
    try:
        element = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located(locator)
        )
        driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});", element)
        logger.debug(f"Scrolled to element: {locator}")
        return True
    except Exception as e:
        logger.warning(f"Could not scroll to {locator} within {timeout}s: {e}")
        return False

def attempt_login(driver):
    """Attempts to log into Prima.it. Returns True if logged in, False otherwise."""
    logger.info("Attempting to log in to Prima.it...")
    
    try:
        driver.get(PRIMA_LOGIN_URL) 
        time.sleep(2) # Give page some time to load, even before explicit waits
    except Exception as e:
        logger.error(f"[KO] Login failed: {e}")
        #driver.quit()
        return False
        
    try:
        if not wait_and_send_keys(driver, LOC_USERNAME_INPUT, PRIMA_USER): return False
        if not wait_and_send_keys(driver, LOC_PASSWORD_INPUT, PRIMA_PASSWORD): return False
        if not wait_and_click(driver, LOC_LOGIN_BUTTON): return False
        
        logger.info("[OK] Successfully logged in.")
        return True
    except Exception as e:
        logger.error(f"[KO] Login failed: {e}")
        #driver.quit()
        return False

def handle_prima_quotation_form(driver, row):
    """
    Navigates the Prima.it quotation form and attempts to get the price.
    Returns the price string or "Not found" or "Error".
    """
    plate = row['Auto']
    logger.info(f"[{plate}] Starting quotation process.")

    try:
        # Step 1: Navigate to new quotation form (assuming already logged in)
        if not wait_and_click(driver, LOC_QUOTATION_BUTTON): raise Exception("Quotation button not found.")
        time.sleep(1) # Small delay for navigation
        if not wait_and_click(driver, LOC_COMPUTE_MOTOR_BUTTON): raise Exception("Compute Motor button not found.")
        time.sleep(1) # Small delay for form load

        # Step 2: Fill initial vehicle/owner data
        if not wait_and_send_keys(driver, LOC_PLATE_INPUT, plate): raise Exception("Plate input not found.")
        
        # Format date as dd/mm/YYYY
        birthday_str = row['Data di nascita'].strftime("%d/%m/%Y")
        if not wait_and_send_keys(driver, LOC_BIRTHDAY_INPUT, birthday_str):
            logger.warning(f"[{plate}] Error on birthday input. Skipping.") # Don't raise, try to continue
        
        scroll_to_element(driver, LOC_PROCEED_BUTTON)
        if not wait_and_click(driver, LOC_PROCEED_BUTTON): raise Exception("Proceed button (page 1) not found.")
        time.sleep(1) # Allow page to transition

        # Step 3: Fill additional owner data (Page 2)
        # Effective Date (optional, only if asked)
        expiration_str = row['Scadenza'].strftime("%d/%m/%Y")
        try:
            effective_date_field = WebDriverWait(driver, 3).until( # Shorter wait for optional field
                EC.element_to_be_clickable(LOC_EFFECTIVE_DATE_INPUT)
            )
            effective_date_field.send_keys(expiration_str)
            logger.info(f"[{plate}] Filled effective date.")
        except:
            logger.info(f"[{plate}] Effective date field not present (all info available).")

        # License Year
        scroll_to_element(driver, LOC_LICENSE_YEAR_DROPDOWN)
        if not wait_and_click(driver, LOC_LICENSE_YEAR_DROPDOWN): raise Exception("License year dropdown not found.")
        time.sleep(0.5) # Short delay for dropdown to open

        licence_year = '' if pd.isna(row['Anno patente']) else str(int(row['Anno patente'])) # Ensure integer for year
        if licence_year:
            if not wait_and_click(driver, LOC_LICENSE_YEAR_OPTION(licence_year)):
                logger.warning(f"[{plate}] Specific license year '{licence_year}' not found. Selecting first available.")
                if not wait_and_click(driver, LOC_DEFAULT_LICENSE_YEAR_OPTION): raise Exception("Default license year option not found.")
        else:
            logger.info(f"[{plate}] License year not provided. Selecting first available.")
            if not wait_and_click(driver, LOC_DEFAULT_LICENSE_YEAR_OPTION): raise Exception("Default license year option not found.")
        time.sleep(0.5) # Delay after selection

        # Residential City
        scroll_to_element(driver, LOC_RESIDENTIAL_CITY_INPUT)
        residential_city = '' if pd.isna(row['Citta di residenza']) else str(row['Citta di residenza'])
        if residential_city:
            if not wait_and_send_keys(driver, LOC_RESIDENTIAL_CITY_INPUT, residential_city): raise Exception("Residential city input not found.")
        else:
            if not wait_and_send_keys(driver, LOC_RESIDENTIAL_CITY_INPUT, 'Torino'): raise Exception("Residential city input not found.")
        time.sleep(1) # Give time for autocomplete
        try:
            if not wait_and_click(driver, LOC_RESIDENTIAL_CITY_SELECT_FIRST_OPTION, timeout=3): # Shorter wait
                logger.warning(f"[{plate}] Residential city autocomplete option not found/clicked.")
        except:
            logger.warning(f"[{plate}] Residential city autocomplete option not present.")

        # CAP (Postal Code)
        cap = '' if pd.isna(row['Cap']) else str(int(row['Cap'])) # Ensure integer for CAP
        if cap:
            if not wait_and_send_keys(driver, LOC_CAP_INPUT, cap): raise Exception("CAP input not found.")
            time.sleep(1) # Give time for autocomplete
            if not wait_and_click(driver, LOC_CAP_SELECT_OPTION(cap), timeout=3):
                logger.warning(f"[{plate}] Specific CAP '{cap}' autocomplete option not found. Attempting default.")
                if not wait_and_send_keys(driver, LOC_CAP_INPUT, '10121'): raise Exception("CAP input not found for default.")
                if not wait_and_click(driver, LOC_DEFAULT_CAP_OPTION, timeout=3): raise Exception("Default CAP option not found.")
        else:
            logger.info(f"[{plate}] CAP not provided. Using default '10121'.")
            if not wait_and_send_keys(driver, LOC_CAP_INPUT, '10121'): raise Exception("CAP input not found for default.")
            time.sleep(1)
            if not wait_and_click(driver, LOC_DEFAULT_CAP_OPTION, timeout=3): raise Exception("Default CAP option not found.")
        
        # Address and Civic Number
        residential_address = '' if pd.isna(row['Indirizzo']) else str(row['Indirizzo'])
        if residential_address:
            if not wait_and_send_keys(driver, LOC_RESIDENTIAL_ADDRESS_INPUT, residential_address): raise Exception("Address input not found.")
        else:
            if not wait_and_send_keys(driver, LOC_RESIDENTIAL_ADDRESS_INPUT, 'Via Roma'): raise Exception("Address input not found.")
        time.sleep(0.5)

        residential_number = '' if pd.isna(row['Civico']) else str(int(row['Civico'])) # Ensure integer for civic number
        if residential_number:
            if not wait_and_send_keys(driver, LOC_RESIDENTIAL_NUMBER_INPUT, residential_number): raise Exception("Civic number input not found.")
        else:
            if not wait_and_send_keys(driver, LOC_RESIDENTIAL_NUMBER_INPUT, '1'): raise Exception("Civic number input not found.")
        time.sleep(0.5)

        # Occupation
        scroll_to_element(driver, LOC_OCCUPATION_DROPDOWN)
        if not wait_and_click(driver, LOC_OCCUPATION_DROPDOWN): raise Exception("Occupation dropdown not found.")
        time.sleep(0.5)
        if not wait_and_click(driver, LOC_OCCUPATION_SELECT_SECOND_OPTION): raise Exception("Occupation option not found.")
        time.sleep(0.5)

        # Civil Status
        scroll_to_element(driver, LOC_CIVIL_STATUS_DROPDOWN)
        if not wait_and_click(driver, LOC_CIVIL_STATUS_DROPDOWN): raise Exception("Civil status dropdown not found.")
        time.sleep(0.5)
        if not wait_and_click(driver, LOC_CIVIL_STATUS_SELECT_FIRST_OPTION): raise Exception("Civil status option not found.")
        time.sleep(0.5)

        # Cell Number
        if not wait_and_send_keys(driver, LOC_CELL_NUMBER_INPUT, '3270692082'): raise Exception("Cell number input not found.")
        time.sleep(0.5)

        # Privacy
        scroll_to_element(driver, LOC_PRIVACY_CHECKBOX)
        if not wait_and_click(driver, LOC_PRIVACY_CHECKBOX): raise Exception("Privacy checkbox not found.")
        time.sleep(0.5)

        # Compute Quotation
        scroll_to_element(driver, LOC_COMPUTE_QUOTATION_BUTTON)
        if not wait_and_click(driver, LOC_COMPUTE_QUOTATION_BUTTON): raise Exception("Compute quotation button not found.")
        logger.info(f"[{plate}] Waiting for quotation results...")
        time.sleep(7) # Longer sleep here as it's a computation

        # Step 4: Extract Quotation Price
        quotation_price_element = WebDriverWait(driver, 15).until( # Increased wait for price
            EC.presence_of_element_located(LOC_QUOTATION_PRICE)
        )
        output = quotation_price_element.text.strip()
        logger.info(f"[{plate}] Found quotation price: '{output}'")
        return output

    except Exception as e:
        logger.error(f"[{plate}] [WARNING] Error during quotation retrieval: {e}")
        # Optionally, save a screenshot for debugging
        # driver.save_screenshot(f"error_{plate}_{datetime.now().strftime('%Y%m%d%H%M%S')}.png")
        return "Error"

# --- Main Automation Logic ---
def main():
    # Read excel with customers data
    try:
        df = pd.read_excel(EXCEL_INPUT_FILE)
        logger.info(f"Loaded {len(df)} records from {EXCEL_INPUT_FILE}")
    except FileNotFoundError:
        logger.error(f"[KO] Input file not found: {EXCEL_INPUT_FILE}. Please ensure it exists.")
        return

    # Prepare DataFrame columns if they don't exist
    if 'Modello' not in df.columns:
        df['Modello'] = None # or pd.NA
    if 'Preventivo' not in df.columns:
        df['Preventivo'] = None # or pd.NA

    # Select range based on date
    start_date = datetime.strptime(START_DATE_STR, "%Y-%m-%d")
    end_date = datetime.strptime(END_DATE_STR, "%Y-%m-%d")

    # Ensure 'Scadenza' and 'Data di nascita' columns are datetime
    df['Scadenza'] = pd.to_datetime(df['Scadenza'], format="%d/%m/%Y", errors='coerce')
    df['Data di nascita'] = pd.to_datetime(df['Data di nascita'], format="%d/%m/%Y", errors='coerce')


    # Filter data: within date range AND 'Preventivo' is null or empty string
    customers_to_process_df = df[
        (df['Scadenza'] >= start_date) &
        (df['Scadenza'] <= end_date) &
        (df['Preventivo'].isnull() | (df['Preventivo'] == ''))
    ].copy() # Use .copy() to avoid SettingWithCopyWarning

    if customers_to_process_df.empty:
        logger.info(f"No records found between {start_date.date()} and {end_date.date()} with missing quotation. Exiting.")
        return

    logger.info(f"[OK] Found {len(customers_to_process_df)} records between {start_date.date()} and {end_date.date()} with missing quotation to process.")
    logger.info("First 5 records to process:")
    logger.info(customers_to_process_df.head())

    # --- Initialize Undetected ChromeDriver ONCE before the loop ---
    logger.info("\nInitializing Undetected ChromeDriver...")
    chrome_options = uc.ChromeOptions()
    # chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-software-rasterizer")
    chrome_options.add_argument("--log-level=3") # Suppress verbose logs from Chrome
    # chrome_options.add_argument("--headless=new") 
    # chrome_options.add_argument("--no-sandbox")
    # chrome_options.add_argument("--disable-dev-shm-usage")

    driver = None
    try:
        driver = uc.Chrome(options=chrome_options)
        logger.info("[OK] Undetected ChromeDriver initialized successfully.")
    except Exception as e:
        logger.critical(f"[KO] Failed to initialize Undetected ChromeDriver: {e}")
        logger.critical("Please ensure Chrome browser is installed and compatible with undetected_chromedriver.")
        return # Exit if driver setup fails


    # Attempt to log in once at the beginning
    if not attempt_login(driver):
        logger.critical("[KO] Initial login failed. Exiting script.")
        #driver.quit()
        return

    # --- Loop through filtered customers to enrich data ---
    processed_count = 0
    for original_idx, row in customers_to_process_df.iterrows():
        plate = row['Auto']
        
        logger.info(f"\n--- Processing plate: {plate} (Original index: {original_idx}) ---")
        '''
        if not driver.service.is_connectable():
            logger.error(f"[KO] Browser session closed unexpectedly during processing. Quitting.")
            break
        '''       
        #Redirect to initial page after the first quotation
        if processed_count > 0:
            driver.get(PRIMA_LOGIN_URL)
            time.sleep(2) # Give the page some time to load after navigating        

        quotation_price = handle_prima_quotation_form(driver, row)

        # Update the ORIGINAL DataFrame 'df' using the original index
        df.at[original_idx, 'Preventivo'] = quotation_price
        
        if quotation_price == "Error":
            logger.warning(f"[{plate}] Quotation price set to 'Error' due to scraping issue.")
        elif quotation_price == "Not found":
            logger.info(f"[{plate}] Quotation price set to 'Not found' from webpage.")
        else:
            logger.info(f"[OK] [{plate}] Quotation price successfully obtained: '{quotation_price}'.")
            processed_count += 1

        time.sleep(random.uniform(3, 6)) # Longer random delay between each customer to appear more human

    # --- Final Save ---
    try:
        # Check if the output file exists and is writable
        if os.path.exists(EXCEL_OUTPUT_FILE) and not os.access(EXCEL_OUTPUT_FILE, os.W_OK):
            logger.error(f"[KO] Output file '{EXCEL_OUTPUT_FILE}' is open or write-protected. Please close it.")
        else:
            df.to_excel(EXCEL_OUTPUT_FILE, index=False)
            logger.info(f"\n[OK] Automation finished. All data processed and updated Excel saved to: {EXCEL_OUTPUT_FILE}")
            logger.info(f"Total quotations successfully retrieved: {processed_count}")
    except Exception as e:
        logger.error(f"\n[KO] Error saving final Excel file: {e}")

    finally:
        # Ensure the driver is quit at the very end
        if driver:
            try:
                #driver.quit()
                logger.info("[OK] Browser closed.")
            except Exception as e:
                logger.error(f"[WARNING] Error closing browser: {e}. This might be an ignored OSError.")

if __name__ == "__main__":
    main()