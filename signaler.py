import time
import requests
from selenium import webdriver
from datetime import datetime as dt
import datetime
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
    wb = Workbook()
    ws = wb.active

    shortdate = dt.today().strftime('%Y-%m-%d')
    date = dt.today().strftime("%H:%M")

    title = 'Stock-picking - Handelsmuligheter'

    ws.title = "Handelsmuligheter"

    headings = ['Overskrift', 'Kandidat', 'Bors', 'Navn', 'Ticker', 'Trendgulv', 'Trendtak','Dato', 'RSI', 'Stock-picking', 'Nivå', 'Signal', 'Risk/Reward', 'Potensial']

    ws.append(headings)

    for index in range(len(data)):
        if(data[index] != None):
            date = data[index]['Date']
            name = data[index]['Name']
            ticker = data[index]['Ticker']
            trendgulv = data[index]['Trendgulv']
            trendtak = data[index]['Trendtak']
            signal = data[index]['Handelsmulighetstype']
            riskreward = data[index]['Riskreward']
            potensial = data[index]['Potensial']
            rsi = data[index]['Rsi']

            ws.append([title, '', '', name, ticker, trendgulv, trendtak, date, rsi, 'handelsmuligheter', 'Kort sikt', signal, riskreward, potensial])

            #print(date, ",", name, ",", ticker, ",", trendgulv, ",", trendtak, ",", signal, ",", riskreward, ",", potensial, ",", rsi)
    
    wb.save("handelsmuligheter-" + str(shortdate) + ".xlsx")


def get_data_from_kortsikt(count, columns):

    #Function to split a string
    splitString = lambda row, line, numb: row.text.splitlines()[line].split()[numb]

    try:
        #Find header
        #header = driver.find_element_by_xpath('//*[@id="ca2017_HeadNameTicker"]').text 
      # Find subheader
        subheader = driver.find_element_by_xpath('//*[@id="ca2017_HeadPriceAndDateInfo"]').text
        
        # Date attributes
        year = subheader.split()[-1]
        month = subheader.split()[-2]
        day = subheader.split()[-3].replace('.', '')

        datetime_object = datetime.datetime.strptime(f'{year} {month} {day}', '%Y %b %d').date()

       #Find the first line for trendgulv and trendtak
        row_1 = driver.find_element_by_xpath('//*[@id="ca2017_QuantIndicatorsTable"]/tbody/tr[1]/td')
      # Find the third line for rsi
        row_2 = driver.find_element_by_xpath('//*[@id="ca2017_QuantIndicatorsTable"]/tbody/tr[3]/td')
       #Get only the ticker from header
        #ticker = header.replace('(','').replace(')','').replace('.', ' ').split()[-2]
       #Get only the trendgulv value from first row
        trendgulv = splitString(row_1, 1, 2)
       #Get only the trendtak value from first row
        trendtak = splitString(row_1, 2, 2)
       #Get RSI
        rsi = splitString(row_2, 1, -1)
       
       #  data_entry = {'dato': dato, 'name': name, 'ticker': ticker, 'signal': indikator, 'kvalitet': kvalitet, 'objektiv': objektiv}

        data = {
            'Date': datetime_object,
            'Name': columns[count]['name'],
            'Ticker': columns[count]['ticker'], 
            'signal': columns[count]['signal'],
            'kvalitet': columns[count]['kvalitet'],
            'objektiv': columns[count]['objektiv'],
            'Trendgulv': trendgulv, 
            'Trendtak': trendtak, 
            'Rsi': rsi,
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
                # Columns
                dato = getColumn(row, 0).text
                name = getColumn(row, 1).text
                ticker = getColumn(row, 2).text
                indikator = getColumn(row, 3).text
                kvalitet = getColumn(row, 4).text
                objektiv = getColumn(row, 5).text

                data_entry = {'dato': dato, 'name': name, 'ticker': ticker, 'signal': indikator, 'kvalitet': kvalitet, 'objektiv': objektiv}

                dataList.append(data_entry)

            except IndexError as e:
                print("error:", e)
    
    return dataList


def get_column_link_from_entries(table, rows):

    # Initialize an empty array to store every link
    links = []
    # Loop over rows inside the table
    for index, row in enumerate(rows):
        # Skip the first line 
        if(index > 0):
            # Find the link and get the 'href' attribute
            link = row.find_elements(By.TAG_NAME, "a")[1].get_attribute('href')
            # Get all the queries inside the link
            queries = parse_qs(urlparse(link).query)
            # Loop over the queries
            for key, value in queries.items():
                # Check if the link has the company id query
                if(key == 'CompanyID'):
                    # Split the string by '&' to get all the individual parts
                    linkSplit = link.split('&')
                    # Change the product parameter
                    linkSplit[1] = 'product=5'
                    # Create a new link 
                    newlink = '&'.join(linkSplit)
                    # Add to a list
                    links.append(newlink)
    return links
        

def signaler():
    # Go to handelsmuligheter table list
    driver.get('https://www.investtech.com/no/market.php?UniverseID=8230&product=16&sgnl=1&indr=0&timeSpan=2')
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
    # Initialize an empty array
    dataList = []

    # Loop through each index within the offset int(offset)
    while index <= int(offset):
        # Go to the next page in universe list
        driver.get(f'https://www.investtech.com/no/market.php?UniverseID=8230&product=16&sgnl=1&indr=0&timeSpan=2&Offset={index}')

        # Find the table by class name
        table = driver.find_element_by_class_name('signalsTable')
        # Get all the rows inside the table
        rows = table.find_elements_by_tag_name('tr')

        if(len(rows) <= 2):
            break

        links = get_column_link_from_entries(table, rows)            
        # Get the data from the table
        columns = get_columns_from_each_entry_in_list(table, rows)

        # Loop through each links
        for count, link in enumerate(links):
            # Go to kort sikt
            driver.get(link)
            # Pause for a moment
            time.sleep(1)
            # Get the data
            data = get_data_from_kortsikt(count, columns)
            # Add the data to the llist
            dataList.append(data)

        time.sleep(1)

        print(columns)

        index += 30
    # Return the data
    return dataList


try: 
    # Set the executable path for chromedriver
    browserPath = 'C:\Program Files (x86)\chromedriver.exe'
    # Start chromedriver
    driver = webdriver.Chrome(browserPath)
    # Start page
    driver.get('https://www.investtech.com/no/market.php?CountryID=1&product=0')
    # Set implicit wait
    driver.implicitly_wait(5)
    # Login
    login()
    # Sleep for a moment before the script continues
    time.sleep(3)

    data = signaler()

    print("data: ", data)

    #export_to_excel(data)

except NoSuchElementException as e:
    print("error: ", e)