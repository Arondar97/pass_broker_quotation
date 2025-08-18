import pandas as pd
from datetime import datetime, timedelta
from selenium import webdriver
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# from webdriver_manager.chrome import ChromeDriverManager # Not strictly needed with undetected_chromedriver
import time
import os # For os.path.exists and future output file handling

# --- Configuration (good practice to define these at the top) ---
EXCEL_INPUT_FILE = "clienti_assicurazioni.xlsx"
EXCEL_OUTPUT_FILE = "customers_updated.xlsx"
START_DATE_STR = "2025-06-01"
END_DATE_STR = "2025-08-01"
INVOICE_URL = "https://intermediari.prima.it/login"
USER = "prima@pass-broker.it"
PASSWORD = "Prima2025!"

# --- Function to extract model (encapsulated for cleanliness) ---
def get_invoice(driver, row):

    plate = row['Auto']
    birthday = row['Data di nascita'].strftime("%d/%m/%Y")
    expiration = row['Scadenza'].strftime("%d/%m/%Y")
    licence_year = '' if pd.isna(row['Anno patente']) else str(row['Anno patente'])
    residential_city = '' if pd.isna(row['Citta di residenza']) else str(row['Citta di residenza'])
    cap = '' if pd.isna(row['Cap']) else str(row['Cap'])
    residential_address = '' if pd.isna(row['Indirizzo']) else str(row['Indirizzo'])
    residential_number = '' if pd.isna(row['Civico']) else str(row['Civico'])

    try:
        driver.get(INVOICE_URL) # Navigate for each plate
        #Log-in
        try: 
            # Handle cookie banner
            user_input = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "id-input-email-id"))
            )
            user_input.send_keys(USER)
            time.sleep(1)

            # Input plate number
            password_input = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "id-input-password-id"))
            )
            password_input.send_keys(PASSWORD)
            time.sleep(1) # Give more time for the page to update after entering plate

            access_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']"))
                )
            access_button.click()
            time.sleep(1)
        except:
            print(f"Already logged in")
        
        quotation_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, ".quotation-link"))
            )
        quotation_button.click()
        time.sleep(1)

        compute_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-test-id='calcola-motor']"))
            )
        compute_button.click()
        time.sleep(1)

        targa = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "plate_number"))
            )
        targa.send_keys(plate)

        try:
            data_nascita = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH,"(//input[@id='owner_birth_date'])[2]"))
                )
            data_nascita.send_keys(birthday)
        except:
            print('Errore sulla data di nascita')

        proceed = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH,"(//button[@type='button'])[1]"))
            )
        driver.execute_script("arguments[0].scrollIntoView({block: 'start'});", proceed)
        proceed.click()
        time.sleep(1)

        #This section will be executed only if data are not enough
        try: 
            effective_day = WebDriverWait(driver, 2).until(
                    EC.element_to_be_clickable((By.XPATH,"(//input[@id='effective_date_date'])[2]"))
                )
            effective_day.send_keys(expiration)
        except:
            print(f"All info available")

        licence_year_field = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR,"div[id='owner_license_year'] span[class='form-select__status']"))
            )
        driver.execute_script("arguments[0].scrollIntoView({block: 'start'});", licence_year_field)
        licence_year_field.click()
        time.sleep(1)     

        if licence_year:
            try:
                licence_year_selection = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH,"(//li[normalize-space()='{licence_year}'])[1]"))
                    )
                licence_year_selection.click()
                time.sleep(1)
            except:
                licence_year_selection = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR,"div[id='owner_license_year'] li:nth-child(1)"))
                    )
                licence_year_selection.click()
                time.sleep(1)               
        else:
            licence_year_selection = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR,"div[id='owner_license_year'] li:nth-child(1)"))
                )
            licence_year_selection.click()
            time.sleep(1)

        residential_city_field = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID,"owner_residential_city"))
            )
        driver.execute_script("arguments[0].scrollIntoView({block: 'start'});", residential_city_field)
        if residential_city:
            residential_city_field.send_keys(residential_city)
        else:
            residential_city_field.send_keys('Torino')
        time.sleep(1)

        try:
            residential_city_select = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR,"div[class='is-valid form-autocomplete is-open is-large is-pristine'] li:nth-child(1)"))
                )
            residential_city_select.click()
            time.sleep(1)
        except:
            print('Not Selected')

        cap_field = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.ID,"owner_residential_cap"))
                )
        if cap:
            cap_field.send_keys(cap)

            cap_selection = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH,"(//li[normalize-space()='{cap}'])[1]"))
                    )
            cap_selection.click()
            time.sleep(1)     
        else:
            cap_field.send_keys('10121')

            cap_selection = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH,"(//li[normalize-space()='10121'])[1]"))
                    )
            cap_selection.click()
            time.sleep(1)           
 
        driver.execute_script("arguments[0].scrollIntoView({block: 'start'});", cap_field)
        cap_field.click()
        time.sleep(1)
        
        residential_address_field = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.ID,"owner_residential_address"))
                )
        if residential_address:
            residential_address_field.send_keys(residential_address)
            time.sleep(1)
        else:
            residential_address_field.send_keys('via Roma')
            time.sleep(1)           
        
        residential_number_field = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.ID,"owner_residential_civic_number"))
                )
        if residential_number:
            residential_number_field.send_keys(residential_number)
            time.sleep(1)
        else:
            residential_number_field.send_keys('1')
            time.sleep(1)            

        occupation = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH,"(//div[@id='owner_occupation'])[1]"))
                )
        driver.execute_script("arguments[0].scrollIntoView({block: 'start'});", occupation)
        occupation.click()
        time.sleep(1)

        occupation_selection = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR,"div[id='owner_occupation'] li:nth-child(2)"))
                )
        occupation_selection.click()
        time.sleep(1)

        civil_status = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH,"(//div[@id='owner_civil_status'])[1]"))
                )
        driver.execute_script("arguments[0].scrollIntoView({block: 'start'});", civil_status)
        civil_status.click()
        time.sleep(1)

        civil_status_selection = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR,"div[id='owner_civil_status'] li:nth-child(1)"))
                )
        civil_status_selection.click()
        time.sleep(1)

        cell_number = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR,"#phone_number"))
                )
        cell_number.send_keys('3270692082')
        time.sleep(1)

        privacy = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR,"label[for='privacy_all']"))
                )
        driver.execute_script("arguments[0].scrollIntoView({block: 'start'});", privacy)
        privacy.click()
        time.sleep(1)

        compute_quotation = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR,".btn.btn--primary[data-test-id='button-calculate-quote']"))
                )
        compute_quotation.click()
        time.sleep(5)

        quotation_price = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR,"div[class='guarantee-box__price guarantee-box__price--highlighted'] span[class='price__value']"))
                )
        content_price = quotation_price.text
        output = content_price.strip()

        if output:
            return output
        else:
            return "Not found"

    except Exception as e:
        print(f"[{plate}] ⚠️ Error during model retrieval: {e}")
        return "Error"

# --- Main Automation Logic ---
def main():
    # Read excel with customers data
    df = pd.read_excel(EXCEL_INPUT_FILE)
    print(f"Loaded {len(df)} records from {EXCEL_INPUT_FILE}")

    # Select range based on date
    start_date = datetime.strptime(START_DATE_STR, "%Y-%m-%d")
    end_date = datetime.strptime(END_DATE_STR, "%Y-%m-%d")

    # Ensure 'Scadenza' column is datetime, coerce errors to NaT
    df['Scadenza'] = pd.to_datetime(df['Scadenza'], format="%d/%m/%Y", errors='coerce')

    # Filter data: within date range AND 'Modello' is null or empty string
    customers_to_process_df = df[
        (df['Scadenza'] >= start_date) &
        (df['Scadenza'] <= end_date) &
        (df['Modello'].isnull() | (df['Modello'] == '')) # Check for both NaN and empty string
    ].copy() # Use .copy() to avoid SettingWithCopyWarning

    if customers_to_process_df.empty:
        print(f"No records found between {start_date.date()} and {end_date.date()} with missing model. Exiting.")
        return

    print(f"✅ Found {len(customers_to_process_df)} records between {start_date.date()} and {end_date.date()} with missing model to process.")
    print("First 5 records to process:")
    print(customers_to_process_df.head())

    # --- Initialize Undetected ChromeDriver ONCE before the loop ---
    print("\nInitializing Undetected ChromeDriver...")
    chrome_options = uc.ChromeOptions()
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-software-rasterizer")
    chrome_options.add_argument("--log-level=3")
    # Uncomment these if you suspect issues related to sandboxing/shared memory on your OS
    # chrome_options.add_argument("--no-sandbox")
    # chrome_options.add_argument("--disable-dev-shm-usage")
    # Optional: run headless (may help with stability, but also detection)
    # chrome_options.add_argument("--headless=new") 

    driver = None # Initialize driver variable to None
    try:
        driver = uc.Chrome(options=chrome_options)
        print("✅ Undetected ChromeDriver initialized successfully.")
    except Exception as e:
        print(f"❌ Failed to initialize Undetected ChromeDriver: {e}")
        print("Please ensure Chrome browser is installed and compatible with undetected_chromedriver.")
        return # Exit if driver setup fails

    # --- Loop through filtered customers to enrich data ---
    processed_count = 0
    for original_idx, row in customers_to_process_df.iterrows(): # Iterate over the filtered DataFrame

        plate = row['Auto']
        
        print(f"\n--- Processing plate: {plate} (Original index: {original_idx}) ---")
        
        price = get_invoice(driver, row)

        df.at[original_idx, 'Preventivo'] = price

        if price == "Error":
            print(f"⚠️ Plate {plate}: Quotation price set to 'Error' due to scraping issue.")
        elif price == "Not found":
            print(f"⚠️ Plate {plate}: Quotation price set to 'Not found' from webpage.")
        else:
            print(f"✅ Plate {plate}: Quotation price set to '{price}'.")
            processed_count += 1

        time.sleep(2) # Small delay between each plate lookup to be less aggressive

    # --- Final Save ---
    try:
        df.to_excel(EXCEL_OUTPUT_FILE, index=False)
        print(f"\n✅ Automation finished. All data processed and updated Excel saved to: {EXCEL_OUTPUT_FILE}")
        print(f"Total models successfully enriched: {processed_count}")
    except Exception as e:
        print(f"\n❌ Error saving final Excel file: {e}")

    finally:
        # Ensure the driver is quit at the very end, outside the loop
        if driver:
            try:
                driver.quit()
                print("✅ Browser closed.")
            except Exception as e:
                print(f"⚠️ Error closing browser: {e}")

if __name__ == "__main__":
    main()