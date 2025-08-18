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

def main():
    chrome_options = uc.ChromeOptions()
    #chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-software-rasterizer")
    chrome_options.add_argument("--log-level=3") # Suppress verbose logs from Chrome

    driver = None #MODIFICA

    driver = uc.Chrome(options=chrome_options)

    driver.get("https://intermediari.prima.it/login")

    element = WebDriverWait(driver, 2).until(
        EC.presence_of_element_located((By.ID, "id-input-email-id"))
    )
    element.send_keys("prima@pass-broker.it")

    element = WebDriverWait(driver, 2).until(
        EC.presence_of_element_located((By.ID, "id-input-password-id"))
    )
    element.send_keys("Prima2025!")

    time.sleep(1)

    element = WebDriverWait(driver, 2).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']"))
    )
    element.click()

    time.sleep(2)

    #driver.quit()

if __name__ == "__main__":
    main()