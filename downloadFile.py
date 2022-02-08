import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException

try: 
    # Launch
    browser_path = 'C:\Program Files (x86)\chromedriver.exe'
    download_path = r'C:\Users\danie\Documents\Programming\teaks'

    option = webdriver.ChromeOptions()
    option.add_argument("--incognito")
    option.add_argument('disable-component-cloud-policy')
    option.add_experimental_option("prefs", {
    "download.prompt_for_download": False,
    "download.directory_upgrade": False,
    "safebrowsing.enabled": True,
    "download.default_directory": download_path
    })

    driver = webdriver.Chrome(executable_path=browser_path, options=option)
    driver.implicitly_wait(10)

    driver.get('https://www.investtech.com/no/market.php?UniverseID=8230&product=0&Offset=100')

    time.sleep(5)

    elem = driver.find_element_by_xpath('//*[@id="login_button"]')
    elem.click()

    time.sleep(5)

    # Login
    username = driver.find_element_by_xpath('//*[@id="login_form"]/div[1]/input')
    password = driver.find_element_by_xpath('//*[@id="login_form"]/div[2]/input')

    username.send_keys('90079450')
    password.send_keys('vinter1234')

    elem = driver.find_element_by_xpath('//*[@id="login_form_submit"]')
    elem.click()

    time.sleep(5)

    # Download Export file from my universe
    driver.get('https://www.investtech.com/no/market.php?uDump=8230')

    time.sleep(5)
        
except NoSuchElementException as e:
    print("error: ", e)
