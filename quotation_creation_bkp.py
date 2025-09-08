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
EXCEL_OUTPUT_FILE = "quotations.xlsx"
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
LOC_COOKIES_BUTTON = (By.CSS_SELECTOR, "button.cookie-policy-accept")
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

# Prima Results
LOC_QUOTATION_PRICE = (By.CSS_SELECTOR, "div[class='guarantee-box__price guarantee-box__price--highlighted'] span[class='price__value']")
LOC_INFORTUNI_PRICE = (By.CSS_SELECTOR, "div[class='guarantee-box guarantee-infortuni_conducente has-bundle-discount-badge'] span[class='price__value']")
LOC_FURTO_DROPDOWN = (By.XPATH, "//div[@class='guarantee-box guarantee-furto_incendio']//div[@class='guarantee-box__optionsWrapper']//div[1]")
LOC_FURTO_SELECT_OPTION = (By.XPATH, "//div[@class='dropdown__option is-open']//li[@class='dropdown__option__list__item'][normalize-space()='Super']")
LOC_FURTO_PRICE = (By.CSS_SELECTOR, "div[class='guarantee-box guarantee-furto_incendio'] span[class='price__value']")
LOC_ASSISTENZA_DROPDOWN = (By.XPATH,"//div[@class='guarantee-box guarantee-assistenza_stradale has-bundle-discount-badge']//div[@class='dropdown__option']")
LOC_ASSISTENZA_SELECT_OPTION = (By.XPATH,"//div[@class='dropdown__option is-open']//li[@class='dropdown__option__list__item'][normalize-space()='Super']")
LOC_ASSISTENZA_PRICE = (By.CSS_SELECTOR, "div[class='guarantee-box guarantee-assistenza_stradale has-bundle-discount-badge'] span[class='price__value']")
LOC_TUTELA_DROPDOWN = (By.XPATH,"//div[@class='guarantee-box guarantee-tutela_legale']//div[@class='dropdown__option']")
LOC_TUTELA_SELECT_OPTION = (By.XPATH,"//li[contains(text(),'Super, fino a € 20.000')]")
LOC_TUTELA_PRICE = (By.CSS_SELECTOR, "div[class='guarantee-box guarantee-tutela_legale'] span[class='price__value']")
LOC_CRISTALLI_DROPDOWN = (By.XPATH,"//div[@class='guarantee-box guarantee-cristalli']//div[@class='dropdown__option']")
LOC_CRISTALLI_SELECT_OPTION = (By.XPATH,"//div[@class='dropdown__option is-open']//li[@class='dropdown__option__list__item'][normalize-space()='Super']")
LOC_CRISTALLI_PRICE = (By.CSS_SELECTOR, "div[class='guarantee-box guarantee-cristalli'] span[class='price__value']")
LOC_EVENTI_DROPDOWN = (By.XPATH,"//div[@class='guarantee-box guarantee-eventi_naturali with-ribbon-badge with-ribbon-badge__border']//div[@class='dropdown__option']")
LOC_EVENTI_SELECT_OPTION = (By.XPATH,"//div[@class='guarantee-box guarantee-eventi_naturali with-ribbon-badge with-ribbon-badge__border']//li[@class='dropdown__option__list__item'][normalize-space()='Super']")
LOC_EVENTI_PRICE = (By.CSS_SELECTOR, "div[class='guarantee-box guarantee-eventi_naturali with-ribbon-badge with-ribbon-badge__border'] span[class='price__value']")
LOC_ATTI_DROPDOWN = (By.XPATH,"//div[@class='guarantee-box guarantee-eventi_sociopolitici with-ribbon-badge with-ribbon-badge__border']//div[@class='dropdown__option']")
LOC_ATTI_SELECT_OPTION = (By.XPATH,"//div[@class='guarantee-box guarantee-eventi_sociopolitici with-ribbon-badge with-ribbon-badge__border']//li[@class='dropdown__option__list__item'][normalize-space()='Super']")
LOC_ATTI_PRICE = (By.CSS_SELECTOR, "div[class='guarantee-box guarantee-eventi_sociopolitici with-ribbon-badge with-ribbon-badge__border'] span[class='price__value']")
LOC_KASKOCOL_DROPDOWN = (By.XPATH,"//div[@class='guarantee-box guarantee-collisione']//div[@class='dropdown__option']")
LOC_KASKOCOL_SELECT_OPTION = (By.XPATH,"//div[@class='guarantee-box guarantee-collisione']//li[@class='dropdown__option__list__item'][normalize-space()='Super']")
LOC_KASKOCOL_PRICE = (By.CSS_SELECTOR, "div[class='guarantee-box guarantee-collisione'] span[class='price__value']") 
LOC_KASKOCOMPL_DROPDOWN = (By.XPATH,"//div[@class='guarantee-box guarantee-kasko with-ribbon-badge with-ribbon-badge__border']//div[@class='dropdown__option']")
LOC_KASKOCOMPL_SELECT_OPTION = (By.XPATH,"//div[@class='guarantee-box guarantee-kasko with-ribbon-badge with-ribbon-badge__border']//li[@class='dropdown__option__list__item'][normalize-space()='Super']")
LOC_KASKOCOMPL_PRICE = (By.CSS_SELECTOR, "div[class='guarantee-box guarantee-kasko with-ribbon-badge with-ribbon-badge__border'] span[class='price__value']")


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
    # Initialize values 
    quotation_price = None
    infortuni_value = None
    furto_value = None
    assistenza_value = None
    tutela_value = None
    cristalli_value = None
    eventi_value = None
    atti_value = None
    kasko_col_value = None
    kasko_compl_value = None

    plate = row['Auto']
    logger.info(f"[{plate}] Starting quotation process.")

    try:
        # Step 1: Navigate to new quotation form (assuming already logged in)
        if not wait_and_click(driver, LOC_QUOTATION_BUTTON, 10): raise Exception("Quotation button not found.")
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
        if not wait_and_click(driver, LOC_PROCEED_BUTTON, 10): raise Exception("Proceed button (page 1) not found.")
        time.sleep(1) # Allow page to transition

        # Step 3: Fill additional owner data (Page 2)
        # Cookies Button (optional, only if asked)
        try:
            effective_date_field = WebDriverWait(driver, 3).until( # Shorter wait for optional field
                EC.element_to_be_clickable(LOC_COOKIES_BUTTON)
            )
            effective_date_field.click()
            logger.info(f"[{plate}] Cookies button clicked")
        except:
            logger.info(f"[{plate}] Cookies button not present (all info available).")

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

        # Extract Quotation Price
        quotation_price_element = WebDriverWait(driver, 15).until( # Increased wait for price
            EC.presence_of_element_located(LOC_QUOTATION_PRICE)
        )
        quotation_price = quotation_price_element.text.strip()

        # Extract Infortuni Price
        scroll_to_element(driver, LOC_INFORTUNI_PRICE)
        infortuni_element = WebDriverWait(driver, 15).until( # Increased wait for price
            EC.presence_of_element_located(LOC_INFORTUNI_PRICE)
        )
        infortuni_value = infortuni_element.text.strip()

        # Extract Furto Price
        scroll_to_element(driver, LOC_FURTO_DROPDOWN)
        if not wait_and_click(driver, LOC_FURTO_DROPDOWN): raise Exception("Furto dropdown not found.")
        time.sleep(0.5)
        if not wait_and_click(driver, LOC_FURTO_SELECT_OPTION): raise Exception("Furto option Super not found.")
        time.sleep(0.5)
        furto_element = WebDriverWait(driver, 15).until( # Increased wait for price
            EC.presence_of_element_located(LOC_FURTO_PRICE)
        )
        furto_value = furto_element.text.strip()

        # Extract Assistenza Price
        scroll_to_element(driver, LOC_ASSISTENZA_DROPDOWN)
        if not wait_and_click(driver, LOC_ASSISTENZA_DROPDOWN): raise Exception("Assistenza dropdown not found.")
        time.sleep(0.5)
        if not wait_and_click(driver, LOC_ASSISTENZA_SELECT_OPTION): raise Exception("Assistenza option Super not found.")
        time.sleep(0.5)
        assistenza_element = WebDriverWait(driver, 15).until( # Increased wait for price
            EC.presence_of_element_located(LOC_ASSISTENZA_PRICE)
        )
        assistenza_value = assistenza_element.text.strip()

        # Extract Tutela Price
        scroll_to_element(driver, LOC_TUTELA_DROPDOWN)
        if not wait_and_click(driver, LOC_TUTELA_DROPDOWN): raise Exception("Tutela dropdown not found.")
        time.sleep(0.5)
        if not wait_and_click(driver, LOC_TUTELA_SELECT_OPTION): raise Exception("Tutela option Super not found.")
        time.sleep(0.5)
        tutela_element = WebDriverWait(driver, 15).until( # Increased wait for price
            EC.presence_of_element_located(LOC_TUTELA_PRICE)
        )
        tutela_value = tutela_element.text.strip()

        # Extract Cristalli Price
        scroll_to_element(driver, LOC_CRISTALLI_DROPDOWN)
        if not wait_and_click(driver, LOC_CRISTALLI_DROPDOWN): raise Exception("Cristalli dropdown not found.")
        time.sleep(0.5)
        if not wait_and_click(driver, LOC_CRISTALLI_SELECT_OPTION): raise Exception("Cristalli option Super not found.")
        time.sleep(0.5)
        cristalli_element = WebDriverWait(driver, 15).until( # Increased wait for price
            EC.presence_of_element_located(LOC_CRISTALLI_PRICE)
        )
        cristalli_value = cristalli_element.text.strip()

        # Extract Eventi Price
        scroll_to_element(driver, LOC_EVENTI_DROPDOWN)
        if not wait_and_click(driver, LOC_EVENTI_DROPDOWN): raise Exception("Eventi dropdown not found.")
        time.sleep(0.5)
        if not wait_and_click(driver, LOC_EVENTI_SELECT_OPTION): raise Exception("Eventi option Super not found.")
        time.sleep(0.5)
        eventi_element = WebDriverWait(driver, 15).until( # Increased wait for price
            EC.presence_of_element_located(LOC_EVENTI_PRICE)
        )
        eventi_value = eventi_element.text.strip()

        # Extract Atti Price
        scroll_to_element(driver, LOC_ATTI_DROPDOWN)
        if not wait_and_click(driver, LOC_ATTI_DROPDOWN): raise Exception("Atti dropdown not found.")
        time.sleep(0.5)
        if not wait_and_click(driver, LOC_ATTI_SELECT_OPTION): raise Exception("Atti option Super not found.")
        time.sleep(0.5)
        atti_element = WebDriverWait(driver, 15).until( # Increased wait for price
            EC.presence_of_element_located(LOC_ATTI_PRICE)
        )
        atti_value = atti_element.text.strip()

        # Extract Kasco Collisioni Price
        scroll_to_element(driver, LOC_KASKOCOL_DROPDOWN)
        if not wait_and_click(driver, LOC_KASKOCOL_DROPDOWN): raise Exception("Kasco Collisioni dropdown not found.")
        time.sleep(0.5)
        if not wait_and_click(driver, LOC_KASKOCOL_SELECT_OPTION): raise Exception("Kasco Collisioni option Super not found.")
        time.sleep(0.5)
        kasko_col_element = WebDriverWait(driver, 15).until( # Increased wait for price
            EC.presence_of_element_located(LOC_KASKOCOL_PRICE)
        )
        kasko_col_value = kasko_col_element.text.strip()

        # Extract Kasco Completo Price
        scroll_to_element(driver, LOC_KASKOCOMPL_DROPDOWN)
        if not wait_and_click(driver, LOC_KASKOCOMPL_DROPDOWN): raise Exception("Kasco Completo dropdown not found.")
        time.sleep(0.5)
        if not wait_and_click(driver, LOC_KASKOCOMPL_SELECT_OPTION): raise Exception("Kasco Completo option Super not found.")
        time.sleep(0.5)
        kasko_compl_element = WebDriverWait(driver, 15).until( # Increased wait for price
            EC.presence_of_element_located(LOC_KASKOCOMPL_PRICE)
        )
        kasko_compl_value = kasko_compl_element.text.strip()

        logger.info(f"[{plate}] Found quotation price: '{quotation_price}'")
        logger.info(f"[{plate}] Found quotation price: '{infortuni_value}'")
        logger.info(f"[{plate}] Found quotation price: '{furto_value}'")
        logger.info(f"[{plate}] Found quotation price: '{assistenza_value}'")
        logger.info(f"[{plate}] Found quotation price: '{tutela_value}'")
        logger.info(f"[{plate}] Found quotation price: '{cristalli_value}'")
        logger.info(f"[{plate}] Found quotation price: '{eventi_value}'")
        logger.info(f"[{plate}] Found quotation price: '{atti_value}'")
        logger.info(f"[{plate}] Found quotation price: '{kasko_col_value}'")
        logger.info(f"[{plate}] Found quotation price: '{kasko_compl_value}'")

        return {
            "Sito": "Prima.it",                # constant or scraped
            "RC": quotation_price,             # already working
            "Infortuni": infortuni_value,      # scrape later
            "Furto_Incendio": furto_value,     # scrape later
            "Assistenza_stradale": assistenza_value,
            "Tutela legale": tutela_value,
            "Cristalli": cristalli_value,
            "Eventi_natuali": eventi_value,
            "Atti_vandalici": atti_value,
            "Kasko_collisione": kasko_col_value,
            "Kasko_completa": kasko_compl_value,
            "Error": "False"
        }

    except Exception as e:
        logger.error(f"[{plate}] [WARNING] Error during quotation retrieval: {e}")
        # Optionally, save a screenshot for debugging
        # driver.save_screenshot(f"error_{plate}_{datetime.now().strftime('%Y%m%d%H%M%S')}.png")
        return {
            "Sito": "Prima.it",                # constant or scraped
            "RC": quotation_price,             # already working
            "Infortuni": infortuni_value,      # scrape later
            "Furto_Incendio": furto_value,     # scrape later
            "Assistenza_stradale": assistenza_value,
            "Tutela legale": tutela_value,
            "Cristalli": cristalli_value,
            "Eventi_natuali": eventi_value,
            "Atti_vandalici": atti_value,
            "Kasko_collisione": kasko_col_value,
            "Kasko_completa": kasko_compl_value,
            "Error": "True"
        }

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

    logger.info("\nInitializing Output dataframe...")
    results_df = pd.DataFrame(columns=["Targa", "Sito", "RC", "Infortuni", "Furto_Incendio", "Assistenza_stradale",
                                       "Tutela legale", "Cristalli", "Eventi_natuali", "Atti_vandalici",
                                        "Kasko_collisione", "Kasko_completa"])

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

        scraped_data = handle_prima_quotation_form(driver, row)

        # Build full record including Targa
        record = {"Targa": row["Auto"], **scraped_data}

        results_df = pd.concat(
            [results_df, pd.DataFrame([record])],
            ignore_index=True
        )
        
        if scraped_data["Error"] == "True":
            logger.warning(f"[{plate}] Error due to scraping issue.")
            processed_count += 1
        elif scraped_data == "Not found":
            logger.info(f"[{plate}] Quotation price set to 'Not found' from webpage.")
            processed_count += 1
        else:
            logger.info(f"[OK] [{plate}] Quotation price successfully obtained: '{scraped_data["RC"]}'.")
            processed_count += 1

        time.sleep(random.uniform(3, 6)) # Longer random delay between each customer to appear more human

    # --- Final Save ---
    try:
        # Check if the output file exists and is writable
        if os.path.exists(EXCEL_OUTPUT_FILE) and not os.access(EXCEL_OUTPUT_FILE, os.W_OK):
            logger.error(f"[KO] Output file '{EXCEL_OUTPUT_FILE}' is open or write-protected. Please close it.")
        else:
            results_df.pop("Error")  
            results_df.to_excel(EXCEL_OUTPUT_FILE, index=False)
            logger.info(f"\n[OK] Automation finished. All data processed and updated Excel saved to: {EXCEL_OUTPUT_FILE}")
    except Exception as e:
        logger.error(f"\n[KO] Error saving final Excel file: {e}")

    finally:
        # Ensure the driver is quit at the very end
        if driver:
            try:
                driver.quit()
                logger.info("[OK] Browser closed.")
            except Exception as e:
                logger.error(f"[WARNING] Error closing browser: {e}. This might be an ignored OSError.")

if __name__ == "__main__":
    main()