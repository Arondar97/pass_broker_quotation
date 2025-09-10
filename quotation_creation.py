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
log_file_path = 'quotation_creation_log.log'
file_handler = logging.FileHandler(log_file_path, mode='w') # 'a' for append mode, w for write
file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

# Create a StreamHandler to print logs to the console
console_handler = logging.StreamHandler()
console_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

logging.basicConfig(level=logging.INFO, # Set overall logging level
                    handlers=[file_handler, console_handler]) # Add both handlers

logger = logging.getLogger(__name__)

# --- Configuration Constants ---
EXCEL_INPUT_FILE = "data_input.xlsx"
EXCEL_OUTPUT_FILE = "quotations.xlsx"
PRIMA_LOGIN_URL = "https://intermediari.prima.it/login"
PRIMA_USER = "prima@pass-broker.it"
PRIMA_PASSWORD = "Prima2025!"

# --- Selenium Locators (Define them once for readability and maintainability) ---
# Login Page
LOC_USERNAME_INPUT = (By.ID, "id-input-email-id")
LOC_PASSWORD_INPUT = (By.ID, "id-input-password-id")
LOC_LOGIN_BUTTON = (By.CSS_SELECTOR, "button[type='submit']")

# Dashboard / Navigation
LOC_ADVERTIZING = (By.CSS_SELECTOR, ".up-x-to-close.userpilot-experience-btn.ref")
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
LOC_INFORTUNI_DISCOUNT = (By.CSS_SELECTOR, "div[data-test-id='infortuni-conducente-bundle-discount-badge'] strong[class='c-brand-dark']")
LOC_FURTO_DROPDOWN = (By.XPATH, "//div[@class='guarantee-box guarantee-furto_incendio']//div[@class='guarantee-box__optionsWrapper']//div[1]")
LOC_FURTO_SELECT_OPTION = (By.XPATH, "//div[@class='dropdown__option is-open']//li[@class='dropdown__option__list__item'][normalize-space()='Super']")
LOC_FURTO_PRICE = (By.CSS_SELECTOR, "div[class='guarantee-box guarantee-furto_incendio'] span[class='price__value']")
LOC_ASSISTENZA_DROPDOWN = (By.XPATH,"//div[@class='guarantee-box guarantee-assistenza_stradale has-bundle-discount-badge']//div[@class='dropdown__option']")
LOC_ASSISTENZA_SELECT_OPTION = (By.XPATH,"//div[@class='dropdown__option is-open']//li[@class='dropdown__option__list__item'][normalize-space()='Super']")
LOC_ASSISTENZA_PRICE = (By.CSS_SELECTOR, "div[class='guarantee-box guarantee-assistenza_stradale has-bundle-discount-badge'] span[class='price__value']")
LOC_ASSISTENZA_DISCOUNT = (By.CSS_SELECTOR, "div[data-test-id='assistenza-stradale-bundle-discount-badge'] strong[class='c-brand-dark']")
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

    plate = row['Targa']
    logger.info(f"[{plate}] Starting quotation process.")

    try:
        #Close any advertizing
        if not wait_and_click(driver, LOC_ADVERTIZING, 3):
            logger.warning("Advertizing not found.")
        
        # Step 1: Navigate to new quotation form (assuming already logged in)
        if not wait_and_click(driver, LOC_QUOTATION_BUTTON, 10):
            logger.warning("First attempt to click quotation button failed. Retrying...")
            time.sleep(2)
            if not wait_and_click(driver, LOC_QUOTATION_BUTTON, 10):
                raise Exception("Quotation button not found after two attempts.")
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
        if not wait_and_click(driver, LOC_PROCEED_BUTTON, 10):
            logger.warning("First attempt to click proceed button failed. Retrying...")
            time.sleep(2)
            if not wait_and_click(driver, LOC_PROCEED_BUTTON, 10):
                raise Exception("Proceed button (page 1) not found.")
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

        # Extract Quotation
        scroll_to_element(driver, LOC_COMPUTE_QUOTATION_BUTTON)
        if not wait_and_click(driver, LOC_COMPUTE_QUOTATION_BUTTON): raise Exception("Compute quotation button not found.")
        logger.info(f"[{plate}] Waiting for quotation results...")
        time.sleep(7) # Longer sleep here as it's a computation

        # Extract Quotation Price
        quotation_price_element = WebDriverWait(driver, 15).until( # Increased wait for price
            EC.presence_of_element_located(LOC_QUOTATION_PRICE)
        )
        quotation_text = quotation_price_element.text.strip()
        quotation_price = float(quotation_text.replace('€', '').replace(',', '.').strip())

        # Extract Infortuni Price
        try:
            scroll_to_element(driver, LOC_INFORTUNI_PRICE)
            infortuni_element = WebDriverWait(driver, 15).until( # Increased wait for price
                EC.presence_of_element_located(LOC_INFORTUNI_PRICE)
            )
            infortuni_text = infortuni_element.text.strip() #estrae solo il test
            infortuni_value = float(infortuni_text.replace('€', '').replace(',', '.').strip())
            try:
                infortuni_discount_element = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located(LOC_INFORTUNI_DISCOUNT)
                )
                infortuni_discount_text = infortuni_discount_element.text.strip()
                infortuni_discount_value = float(infortuni_discount_text.replace('€', '').replace(',', '.').strip())
            except: 
                logger.warning(f"[{plate}] Nessuno sconto sulla Garanzia")
                infortuni_discount_value = 0

            infortuni_value -= infortuni_discount_value #Valore finale
        except:
            logger.warning(f"[{plate}] Garanzia infortuni non è presente.")

        # Extract Furto Price
        try:
            scroll_to_element(driver, LOC_FURTO_DROPDOWN)
            if not wait_and_click(driver, LOC_FURTO_DROPDOWN):
                time.sleep(0.2)
            if not wait_and_click(driver, LOC_FURTO_SELECT_OPTION):
                time.sleep(0.2)
            furto_element = WebDriverWait(driver, 15).until( # Increased wait for price
                EC.presence_of_element_located(LOC_FURTO_PRICE)
            )
            furto_text = furto_element.text.strip()
            furto_value = float(furto_text.replace('€', '').replace(',', '.').strip())
        except:
            logger.warning(f"[{plate}] Garanzia furto non è presente.")

        # Extract Assistenza Price
        try:        
            scroll_to_element(driver, LOC_ASSISTENZA_DROPDOWN)
            if not wait_and_click(driver, LOC_ASSISTENZA_DROPDOWN):
                time.sleep(0.2)
            if not wait_and_click(driver, LOC_ASSISTENZA_SELECT_OPTION):
                time.sleep(0.2)
            assistenza_element = WebDriverWait(driver, 15).until( # Increased wait for price
                EC.presence_of_element_located(LOC_ASSISTENZA_PRICE)
            )
            assistenza_text = assistenza_element.text.strip()
            assistenza_value = float(assistenza_text.replace('€', '').replace(',', '.').strip())
            try:
                assistenza_discount_element = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located(LOC_ASSISTENZA_DISCOUNT)
                )
                assistenza_discount_text = assistenza_discount_element.text.strip()
                assistenza_discount_value = float(assistenza_discount_text.replace('€', '').replace(',', '.').strip())
            except:
                logger.warning(f"[{plate}] Nessuno sconto sull'assistenza")
                assistenza_discount_value = 0
            assistenza_value -= assistenza_discount_value
        except:
            logger.warning(f"[{plate}] Garanzia assistenza non è presente.")

        # Extract Tutela Price
        try:
            scroll_to_element(driver, LOC_TUTELA_DROPDOWN)
            if not wait_and_click(driver, LOC_TUTELA_DROPDOWN):
                time.sleep(0.2)
            if not wait_and_click(driver, LOC_TUTELA_SELECT_OPTION):
                time.sleep(0.2)
            tutela_element = WebDriverWait(driver, 15).until( # Increased wait for price
                EC.presence_of_element_located(LOC_TUTELA_PRICE)
            )
            tutela_text = tutela_element.text.strip()
            tutela_value = float(tutela_text.replace('€', '').replace(',', '.').strip())
        except:
            logger.warning(f"[{plate}] Garanzia tutela non è presente.")

        # Extract Cristalli Price
        try:
            scroll_to_element(driver, LOC_CRISTALLI_DROPDOWN)
            if not wait_and_click(driver, LOC_CRISTALLI_DROPDOWN):
                time.sleep(0.2)
            if not wait_and_click(driver, LOC_CRISTALLI_SELECT_OPTION):
                time.sleep(0.2)
            cristalli_element = WebDriverWait(driver, 15).until( # Increased wait for price
                EC.presence_of_element_located(LOC_CRISTALLI_PRICE)
            )
            cristalli_text = cristalli_element.text.strip()
            cristalli_value = float(cristalli_text.replace('€', '').replace(',', '.').strip())
        except:
            logger.warning(f"[{plate}] Garanzia cristalli non è presente.")

        # Extract Eventi Price
        try:
            scroll_to_element(driver, LOC_EVENTI_DROPDOWN)
            if not wait_and_click(driver, LOC_EVENTI_DROPDOWN):
                time.sleep(0.2)
            if not wait_and_click(driver, LOC_EVENTI_SELECT_OPTION):
                time.sleep(0.2)
            eventi_element = WebDriverWait(driver, 15).until( # Increased wait for price
                EC.presence_of_element_located(LOC_EVENTI_PRICE)
            )
            eventi_text = eventi_element.text.strip()
            eventi_value = float(eventi_text.replace('€', '').replace(',', '.').strip())
        except:
            logger.warning(f"[{plate}] Eventi non è presente.")
        
        # Extract Atti Price
        try:
            scroll_to_element(driver, LOC_ATTI_DROPDOWN)
            if not wait_and_click(driver, LOC_ATTI_DROPDOWN):
                time.sleep(0.2)
            if not wait_and_click(driver, LOC_ATTI_SELECT_OPTION):
                time.sleep(0.2)
            atti_element = WebDriverWait(driver, 15).until( # Increased wait for price
                EC.presence_of_element_located(LOC_ATTI_PRICE)
            )
            atti_text = atti_element.text.strip()
            atti_value = float(atti_text.replace('€', '').replace(',', '.').strip())
        except:
            logger.warning(f"[{plate}] Atti non è presente.")        

        # Extract Kasco Collisioni Price
        try:
            scroll_to_element(driver, LOC_KASKOCOL_DROPDOWN)
            if not wait_and_click(driver, LOC_KASKOCOL_DROPDOWN):
                time.sleep(0.2)
            if not wait_and_click(driver, LOC_KASKOCOL_SELECT_OPTION):
                time.sleep(0.2)
            kasko_col_element = WebDriverWait(driver, 15).until( # Increased wait for price
                EC.presence_of_element_located(LOC_KASKOCOL_PRICE)
            )
            kasko_col_text = kasko_col_element.text.strip()
            kasko_col_value = float(kasko_col_text.replace('€', '').replace(',', '.').strip())
        except:
            logger.warning(f"[{plate}] Kasco Collisioni non è presente.")                

        # Extract Kasco Completo Price
        try:
            scroll_to_element(driver, LOC_KASKOCOMPL_DROPDOWN)
            if not wait_and_click(driver, LOC_KASKOCOMPL_DROPDOWN):
                time.sleep(0.5)
            if not wait_and_click(driver, LOC_KASKOCOMPL_SELECT_OPTION):
                time.sleep(0.5)
            kasko_compl_element = WebDriverWait(driver, 15).until( # Increased wait for price
                EC.presence_of_element_located(LOC_KASKOCOMPL_PRICE)
            )
            kasko_compl_text = kasko_compl_element.text.strip()
            kasko_compl_value = float(kasko_compl_text.replace('€', '').replace(',', '.').strip())
        except:
            logger.warning(f"[{plate}] Kasco Completo non è presente.") 

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

def upsert_excel_data(dataframe_to_save, excel_filepath):
    try:
        # Load the existing data, or create an empty DataFrame if the file doesn't exist
        if os.path.exists(excel_filepath):
            existing_df = pd.read_excel(excel_filepath, dtype={'Targa': str})
        else:
            existing_df = pd.DataFrame()
        # Read the input data excel
        existing_df = pd.read_excel(excel_filepath)
        
        # Concatenate the existing data with the new data.
        # Set the key column ('Targa') as the index for the merge.
        combined_df = pd.concat([
            existing_df.set_index('Targa'), 
            dataframe_to_save.set_index('Targa')
        ])
        # Drop duplicates based on the index, keeping the *last* (newest) row.
        # This effectively updates existing rows.
        combined_df = combined_df[~combined_df.index.duplicated(keep='last')]
        
        # Reset the index to turn the 'Auto' column back into a regular column
        combined_df.reset_index(inplace=True)

        # Save the combined DataFrame back to the Excel file
        combined_df.to_excel(excel_filepath, index=False)
        logging.info(f"[OK] Excel file {excel_filepath} successfully updated.")        

    except FileNotFoundError:
        logger.warning(f"[ATTENZIONE] Excel file '{excel_filepath}' not found. Creating a new one.")
        # If the file doesn't exist, just save the new DataFrame directly.
        dataframe_to_save.to_excel(excel_filepath, index=False)
        logger.info(f"[OK] New Excel file created with the data.")
    except Exception as e:
        logger.error(f"[KO] An error occurred saving data into {excel_filepath}: {e}")

# --- Main Automation Logic ---
def run_quotation_process(df=None):

    if df is None:
        df = pd.read_excel(EXCEL_INPUT_FILE)

    # 1. Add "Processato" column and save all raw data to a new input file
    if 'Processato' not in df.columns:
        df['Processato'] = 'NO'

    df['Data_inserimento'] = datetime.now().date()

    if 'Preventivo' not in df.columns:
        df['Preventivo'] = None # or pd.NA

    customers_to_process_df = df[
        (df['Processato'] == 'NO')
    ]

    # Ensure 'Scadenza' and 'Data di nascita' columns are datetime
    customers_to_process_df['Scadenza'] = pd.to_datetime(customers_to_process_df['Scadenza'], format="%d/%m/%Y", errors='coerce')
    customers_to_process_df['Data di nascita'] = pd.to_datetime(customers_to_process_df['Data di nascita'], format="%d/%m/%Y", errors='coerce')

    if customers_to_process_df.empty:
        return "I dati non sono stati inseriti correttamente nell'Excel. Riprovare."

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
        driver = uc.Chrome(version_main=139, options=chrome_options)
        logger.info("[OK] Undetected ChromeDriver initialized successfully.")
    except Exception as e:
        logger.critical(f"[KO] Failed to initialize Undetected ChromeDriver: {e}")
        logger.critical("Please ensure Chrome browser is installed and compatible with undetected_chromedriver.")
        return # Exit if driver setup fails

    results = []
    # Attempt to log in once at the beginning
    if not attempt_login(driver):
        logger.critical("[KO] Initial login failed. Exiting script.")
        driver.quit()
        return "Errore: Fallimento nell'inizializzazione del browser o nel login."

    # --- Loop through filtered customers to enrich data ---
    processed_count = 0

    for original_idx, row in customers_to_process_df.iterrows():
        plate = row['Targa']
        
        logger.info(f"\n--- Processing plate: {plate} (Original index: {original_idx}) ---")
      
        #Redirect to initial page after the first quotation
        if processed_count > 0:
            driver.get(PRIMA_LOGIN_URL)
            time.sleep(2) # Give the page some time to load after navigating        
        scraped_data = handle_prima_quotation_form(driver, row)
        # Build full record including Targa
        if scraped_data["Error"] != "True":
            record = {"Targa": row["Targa"], **scraped_data}
            results_df = pd.concat(
                [results_df, pd.DataFrame([record])],
                ignore_index=True
            )
        
        if scraped_data["Error"] == "True":
            logger.warning(f"[KO] [{plate}] Error due to scraping issue.")
            results.append(f"[KO] [{plate}] Error due to scraping issue.")
            processed_count += 1
        elif scraped_data == "Not found":
            logger.info(f"[{plate}] Quotation price set to 'Not found' from webpage.")
            results.append(f"[{plate}] Quotation price set to 'Not found' from webpage.")
            processed_count += 1
        else:
            logger.info(f"[OK] [{plate}] Quotation price successfully obtained: '{scraped_data["RC"]}'.")
            results.append(f"[OK] [{plate}] Quotation price successfully obtained: '{scraped_data["RC"]}'.")
            customers_to_process_df.at[original_idx, 'Processato'] = 'SI'
            processed_count += 1
        time.sleep(random.uniform(3, 6)) # Longer random delay between each customer to appear more human
    # --- Final Save ---
    
    # Save input data
    upsert_excel_data(customers_to_process_df,EXCEL_INPUT_FILE)

    # Save output data
    try:
        # Check if the output file exists and is writable
        if os.path.exists(EXCEL_OUTPUT_FILE) and not os.access(EXCEL_OUTPUT_FILE, os.W_OK):
            logger.error(f"[KO] Output file '{EXCEL_OUTPUT_FILE}' is open or write-protected. Please close it.")
            results.append(f"[KO] Output file '{EXCEL_OUTPUT_FILE}' is open or write-protected. Please close it.")
        else:
            results_df.pop("Error")
            results_df["Data_inserimento"] = datetime.now()
            if not results_df.empty:
                upsert_excel_data(results_df,EXCEL_OUTPUT_FILE)
                logger.info(f"\n[OK] Automation finished. All data processed and updated Excel saved to: {EXCEL_OUTPUT_FILE}")
                results.append(f"\n[OK] Automation finished. All data processed and updated Excel saved to: {EXCEL_OUTPUT_FILE}")
            else:
                logger.info(f"\nNessun preventivo calcolato.")
    except Exception as e:
        logger.error(f"\n[KO] Error saving final Excel file: {e}")
        results.append(f"\n[KO] Error saving final Excel file: {e}")

    finally:
        # Ensure the driver is quit at the very end
        if driver:
            try:
                driver.quit()
                logger.info("[OK] Browser closed.")
            except Exception as e:
                logger.error(f"[WARNING] Error closing browser: {e}. This might be an ignored OSError.")
    
    return "\n".join(results)