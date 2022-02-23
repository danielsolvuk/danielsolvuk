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
from openpyxl import Workbook, load_workbook


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


def export_to_excel(data):
    print(data)
    # wb = Workbook()
    # ws = wb.active

    # shortdate = datetime.today().strftime('%Y-%m-%d')
    # date = datetime.today().strftime("%H:%M")

    # title = 'Stock-picking - Handelsmuligheter'

    # ws.title = "Handelsmuligheter"

    # headings = [
    #     'Overskrift'
    #     'Kandidat'
    #     'Bors', 
    #     'Navn', 
    #     'Ticker', 
    #     'Trendgulv', 
    #     'Trendtak', 
    #     'Risk/Reward', 
    #     'Potensial', 
    #     'RSI'
    # ]

    # # ws.append(headings)

    # # for key in data:
    # #     for value in data[key]:
    # #         ws.append([title, '', value['name'], value['ticker'], value['trendgulv'], value['trendtak']])

    # # wb.save("handelsmuligheter-" + str(shortdate) + ".xlsx")

def get_data_from_kortsikt():

    #Function to split a string
    splitString = lambda row, line, numb: row.text.splitlines()[line].split()[numb]

    try:
        #Find header
        header = driver.find_element_by_xpath('//*[@id="ca2017_HeadNameTicker"]').text 
      # Find subheader
        subheader = driver.find_element_by_xpath('//*[@id="ca2017_HeadPriceAndDateInfo"]').text
       #Date attributes
        year = subheader.split()[-1]
        month = subheader.split()[-2]
        day = subheader.split()[-3].replace('.', '')

        datetime_object = datetime.strptime(f'{year} {month} {day}', '%Y %b %d').date()

       #Find the first line for trendgulv and trendtak
        row_1 = driver.find_element_by_xpath('//*[@id="ca2017_QuantIndicatorsTable"]/tbody/tr[1]/td')
      # Find the third line for rsi
        row_2 = driver.find_element_by_xpath('//*[@id="ca2017_QuantIndicatorsTable"]/tbody/tr[3]/td')
       #Get only the ticker from header
        ticker = header.replace('(','').replace(')','').replace('.', ' ').split()[-2]
       #Get only the trendgulv value from first row
        trendgulv = splitString(row_1, 1, 2)
       #Get only the trendtak value from first row
        trendtak = splitString(row_1, 2, 2)
       #Get RSI
        rsi = splitString(row_2, 1, -1)

        data = {
            'Date': datetime_object,
            'Ticker': ticker, 
            'Trendgulv': trendgulv, 
            'Trendtak:': trendtak, 
            'Rsi:': rsi
        } 

        return data

    except Exception as e:
        print("error", e)


def get_columns_from_each_entry_in_list(table, rows):
    # Initialize an empty array
    dataList = []

    getColumn = lambda row, number: row.find_elements(By.TAG_NAME, "td")[number] 

    # Loop through each row

    for index, row in enumerate(rows):
        if(index > 0):
            try:
                # Get only the first link from each entry 
                riskreward = getColumn(row, 4).text
                potensial = getColumn(row, 5).text

                data_entry = {'riskreward': riskreward, 'potensial': potensial}

                dataList.append(data_entry)

            except IndexError:
                dataList.append({'null'})
    
    return dataList


def get_column_link_from_entries(table, rows):

    links = []
    
    for index, row in enumerate(rows):
        if(index > 0):
            link = row.find_elements(By.TAG_NAME, "a")[1].get_attribute('href')

            queries = parse_qs(urlparse(link).query)

            for key, value in queries.items():
                if(key == 'CompanyID'):
                    linkSplit = link.split('&')
                    linkSplit[1] = 'product=5'
                
                    newlink = '&'.join(linkSplit)

                    links.append(newlink)
    return links
        

def handelsmuligheter():
    # Go to handelsmuligheter table list
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

        # Find the table by class name
        table = driver.find_element_by_class_name('tradingOpportunitiesListView')
        # Get all the rows inside the table
        rows = table.find_elements_by_tag_name('tr')

        links = get_column_link_from_entries(table, rows)
        columns = get_columns_from_each_entry_in_list(table, rows)

        for link in links:
            driver.get(link)
            kortsikt = get_data_from_kortsikt()

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

    handelsmuligheter()

    # Pause for a moment
    time.sleep(3)

except NoSuchElementException as e:
    print("error: ", e)
