from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException, NoAlertPresentException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from config import *
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
import logging

if __name__ == '__main__': 
    
    options = webdriver.ChromeOptions()
    options.add_argument('--user-data-dir='+CHROME_PROFILE_PATH)
    options.add_argument('--profile-directory=Default')

    delay = 3
    
    service = Service('./chromedriver.exe')
    chrome_browser = webdriver.Chrome(service=service, options=options)
    
    data = pd.read_excel(filepath = EXCEL_PATH, engine='openpyxl', sheet_name=SHEETS)[CURRENT_SHEET]
    chrome_browser.get('https://web.whatsapp.com')

    try:
        WebDriverWait(chrome_browser, delay).until(EC.presence_of_element_located((By.ID, 'waitCreate')))
    except TimeoutException:
        logging.exception("Loading took too much time")

    for index, i in data.iterrows():
        chrome_browser.get('https://web.whatsapp.com/send?phone=+'+str(i['Phone Number'])+'&text='+i['Message'])
        time.sleep(4)
        
        try:
            send_button = chrome_browser.find_element_by_xpath(SEND_BUTTON_XPATH)
            send_button.click()
            time.sleep(2)
        except NoSuchElementException:
            logging.exception('Button class not found')
            pass     
           
        try:
            chrome_browser.switch_to.alert.accept()
        except NoAlertPresentException:
            logging.exception('Alert not found')
            pass
        
    chrome_browser.quit()
        

    
    
    
    