import time
import requests
from selenium import webdriver
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from urllib.parse import parse_qs, urlparse


def login():
    # Click on login button
    elem = driver.find_element_by_xpath('//*[@id="login_button"]')
    elem.click()
    # Login
    username = driver.find_element_by_xpath('//*[@id="login_form"]/div[1]/input')
    password = driver.find_element_by_xpath('//*[@id="login_form"]/div[2]/input')
    # Send keys
    username.send_keys('90079450')
    password.send_keys('vinter1234')
    # Click on login button
    elem = driver.find_element_by_xpath('//*[@id="login_form_submit"]')
    elem.click()



def export_to_excel(arr):
    print(arr)

def get_data_from_kortsikt(link):

    driver.get(link)

    splitString = lambda row, line, numb: row.text.splitlines()[line].split()[numb]

    try:
        # Find header
        header = driver.find_element_by_xpath('//*[@id="ca2017_HeadNameTicker"]').text 
        # Find subheader
        subheader = driver.find_element_by_xpath('//*[@id="ca2017_HeadPriceAndDateInfo"]').text

        year = subheader.split()[-1]
        month = subheader.split()[-2]
        day = subheader.split()[-3].replace('.', '')

        datetime_object = datetime.strptime(f'{year} {month} {day}', '%Y %b %d').date()

        # Find the first line for trendgulv and trendtak
        row_1 = driver.find_element_by_xpath('//*[@id="ca2017_QuantIndicatorsTable"]/tbody/tr[1]/td')
        # Find the third line for rsi
        row_2 = driver.find_element_by_xpath('//*[@id="ca2017_QuantIndicatorsTable"]/tbody/tr[3]/td')
        # Get only the ticker from header
        ticker = header.replace('(','').replace(')','').replace('.', ' ').split()[-2]
        # Get only the trendgulv value from first row
        trendgulv = splitString(row_1, 1, 2)
        # Get only the trendtak value from first row
        trendtak = splitString(row_1, 2, 2)
        # Get RSI
        rsi = splitString(row_2, 1, -1)

        print(datetime_object)

        data = {
            'Date': datetime_object,
            'Ticker': ticker, 
            'Trendgulv': trendgulv, 
            'Trendtak:': trendtak, 
            'Rsi:': rsi
        } 

        export_to_excel([data])
        
        # Return as dictionary
        #return {'Trendgulv': trendgulv, 'Trendtak': trendtak, 'Rsi': rsi}

    except Exception as e:
        print("error", e)


def go_to_each_stockpage_from_universeList(links):
    for link in links:
        queries = parse_qs(urlparse(link).query)
        for key, value in queries.items():
            if(key == 'CompanyID'):
                print(key, ':', value)
    

        # queries = parse_qs(urlparse(link).query)
        # for key, value in queries.items():
        #     if(key == 'product'):
        #         if('5' in value):
        #             driver.get(link)
                    #get_data_from_kortsikt(link)


def get_each_stockpage_from_universeList():
    table = driver.find_element_by_class_name('tradingOpportunitiesListView')
    rows = table.find_elements_by_tag_name('tr')
    
    links = []

    for row in rows:
        link = row.find_elements(By.TAG_NAME, "a")[1]
        links.append(link.get_attribute('href'))

    go_to_each_stockpage_from_universeList(links)
    
#//*[@id="productBox"]/table[4]/tbody/tr[32]/td/a[1]
#//*[@id="productBox"]/table[4]/tbody/tr[32]/td/a[2]
def go_to_universeList():
    
    driver.get('https://www.investtech.com/no/market.php?UniverseID=8230&product=53&toRpt=0&sgnl=1&timeSpan=2')
    # Find last page element
    last_page = driver.find_element_by_xpath('//*[@id="productBox"]/table[4]/tbody/tr[32]/td/a[2]')
    # Retrieve the href
    last_page_href = last_page.get_attribute('href')
    # Get params from url
    query_params = parse_qs(urlparse(last_page_href).query)
    # Get the offset value
    offset = query_params['Offset'][0]
    # Start at index = 0
    index = 0

    # Loop through each index within the offset int(offset)
    while index <= int(offset):
        # Go to the next page in universe list
        driver.get(f'https://www.investtech.com/no/market.php?UniverseID=8230&product=53&toRpt=0&sgnl=1&timeSpan=2&Offset={index}')

        get_each_stockpage_from_universeList()

        index += 30


try: 
    # Launch
    browserPath = 'C:\Program Files (x86)\chromedriver.exe'
    driver = webdriver.Chrome(browserPath)
    driver.get('https://www.investtech.com/no/market.php?CountryID=1&product=0')
    # Set implicit wait
    driver.implicitly_wait(10)

    login()

    time.sleep(5)

    go_to_universeList()

    # Pause for a moment
    time.sleep(3)

except NoSuchElementException as e:
    print("error: ", e)
