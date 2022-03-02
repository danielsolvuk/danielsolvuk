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
    # Start the workbook
    wb = Workbook()
    # Get the active sheet
    ws = wb.active

    # Date attributes
    shortdate = dt.today().strftime('%Y-%m-%d')
    date = dt.today().strftime("%H:%M")

    # The title
    title = 'Stock-picking - Handelsmuligheter'
    # Sheet title
    ws.title = "Handelsmuligheter"
    # Headings 
    headings = ['Overskrift', 'Kandidat', 'Bors', 'Navn', 'Ticker', 'Trendgulv', 'Trendtak','Dato', 'RSI', 'Stock-picking', 'Nivå', 'Signal', 'Risk/Reward', 'Potensial']
    # add the headings to the active sheet
    ws.append(headings)
    # Loop over the data
    for index in range(len(data)):
        # Check if list has 'None' values and skip the bad lines
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

            # Add the data to the sheet
            ws.append([title, '', '', name, ticker, trendgulv, trendtak, date, rsi, 'handelsmuligheter', 'Kort sikt', signal, riskreward, potensial])
    
    # Save the excel file
    wb.save("handelsmuligheter-" + str(shortdate) + ".xlsx")



def get_data_from_subheader():

    try:
        
        #Find header
        header = driver.find_element_by_xpath('//*[@id="ca2017_HeadNameTicker"]').text 
        
        #Find subheader
        subheader = driver.find_element_by_xpath('//*[@id="ca2017_HeadPriceAndDateInfo"]').text
        
        #   Date attributes
        year = subheader.split()[-1]
        month = subheader.split()[-2]
        day = subheader.split()[-3].replace('.', '')
        
        datetime_object = datetime.datetime.strptime(f'{year} {month} {day}', '%Y %b %d').date()

        ticker = header.replace('(','').replace(')','').replace('.', ' ').split()[-2]
        
        data = ( datetime_object, ticker )

    except Exception as e:
        data = ('Null', 'Null')

    return list(data)

def get_data_from_row_1(splitString):
    try:
        row_1 = driver.find_element_by_xpath('//*[@id="ca2017_QuantIndicatorsTable"]/tbody/tr[1]/td')
        trendgulv = splitString(row_1, 1, 2)
        trendtak = splitString(row_1, 2, 2)

        data = ( trendgulv, trendtak )

    except Exception as e:
        data = ( 'Null', 'Null' )

    return list(data)


def get_data_from_row_2(splitString):
    
    try:
        row_2 = driver.find_element_by_xpath('//*[@id="ca2017_QuantIndicatorsTable"]/tbody/tr[3]/td')
        rsi = splitString(row_2, 1, -1)

        value = rsi

    except Exception as e:
        value = 'Not found'

    return value

def get_data_from_kortsikt(count, columns):
    #   Function to split a string
    splitString = lambda row, line, numb: row.text.splitlines()[line].split()[numb]

    datetime_obj = get_data_from_subheader()[0]
    ticker = get_data_from_subheader()[1]

    trendgulv = get_data_from_row_1(splitString)[0]
    trendtak = get_data_from_row_1(splitString)[1]

    rsi = get_data_from_row_2(splitString)

    print(datetime_obj, trendgulv, trendtak, rsi)

    # return  {
    #     'Date': datetime_obj,
    #     'Name': columns[count]['name'],
    #     'Ticker': columns[count]['ticker'], 
    #     'Days_since': columns[count]['dager_siden'],
    #     'Pivotpunkt': columns[count]['pivotpunkt'],
    #     'Days_since_id': columns[count]['dager_siden_id'],
    #     'Signal': columns[count]['signal'],
    #     'Trendgulv': trendgulv, 
    #     'Trendtak': trendtak, 
    #     'Rsi': rsi,
    # } 

    return {
            'Date': datetime_obj,
            'Name': columns[count]['name'],
            'Ticker': ticker, 
            'Trendgulv': trendgulv, 
            'Trendtak': trendtak, 
            'Rsi': rsi,
            'Handelsmulighetstype': columns[count]['handelstype'],
            'Riskreward': columns[count]['riskreward'],
            'Potensial': columns[count]['potensial']
        } 

# def get_data_from_kortsikt(count, columns):

#     # Function to split a string
#     splitString = lambda row, line, numb: row.text.splitlines()[line].split()[numb]

#     try:
#         # Find header
#         header = driver.find_element_by_xpath('//*[@id="ca2017_HeadNameTicker"]').text 
        
#         # Find subheader
#         subheader = driver.find_element_by_xpath('//*[@id="ca2017_HeadPriceAndDateInfo"]').text
        
#         # Date attributes
#         year = subheader.split()[-1]
#         month = subheader.split()[-2]
#         day = subheader.split()[-3].replace('.', '')

#         # Create date from the date attributes
#         datetime_object = datetime.datetime.strptime(f'{year} {month} {day}', '%Y %b %d').date()

#         # Find the first line for trendgulv and trendtak
#         row_1 = driver.find_element_by_xpath('//*[@id="ca2017_QuantIndicatorsTable"]/tbody/tr[1]/td')
#         # Find the third line for rsi
#         row_2 = driver.find_element_by_xpath('//*[@id="ca2017_QuantIndicatorsTable"]/tbody/tr[3]/td')
#         # Get only the ticker from header
#         ticker = header.replace('(','').replace(')','').replace('.', ' ').split()[-2]
#         # Get only the trendgulv value from first row
#         trendgulv = splitString(row_1, 1, 2)
#         # Get only the trendtak value from first row
#         trendtak = splitString(row_1, 2, 2)
#         # Get RSI
#         rsi = splitString(row_2, 1, -1)

#         # Put the data together 
#         data = {
#             'Date': datetime_object,
#             'Name': columns[count]['name'],
#             'Ticker': ticker, 
#             'Trendgulv': trendgulv, 
#             'Trendtak': trendtak, 
#             'Rsi': rsi,
#             'Handelsmulighetstype': columns[count]['handelstype'],
#             'Riskreward': columns[count]['riskreward'],
#             'Potensial': columns[count]['potensial']
#         } 

#         return data

#     except Exception as e:
#         print("error", e)


def get_columns_from_each_entry_in_list(table, rows):
    # Initialize an empty array to store each entry from the table
    dataList = []
    # A function to get the data inside the columns 
    getColumn = lambda row, number: row.find_elements(By.TAG_NAME, "td")[number] 
    # Loop through each row
    for index, row in enumerate(rows):
        # Skip the first line which is the header
        if(index > 0):
            try:
                # Get the first column (name) and the text value
                name = getColumn(row, 0).text
                # Get the second column (handelsmulighetstype) and the text value
                handelstype = getColumn(row, 2).text
                # Get the third column (riskreward) and the text value
                riskreward = getColumn(row, 4).text
                # Get the fourth column (potensial) and the text value
                potensial = getColumn(row, 5).text
                # Put each column inside a dictionary 
                data_entry = {'name': name, 'handelstype': handelstype, 'riskreward': riskreward, 'potensial': potensial}
                # Append the dictionary inside a list to get all the entries
                dataList.append(data_entry)
            
            # Skip bad lines, don't stop the script
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
    # Initialize an empty array
    dataList = []

    # Loop through each index within the offset int(offset)
    while index <= int(offset):
        # Go to the next page in universe list
        driver.get(f'https://www.investtech.com/no/market.php?UniverseID=8230&product=53&toRpt=0&sgnl=1&timeSpan=2&Offset={index}')
        # Find the table by class name
        table = driver.find_element_by_class_name('tradingOpportunitiesListView')
        # Get all the rows inside the table
        rows = table.find_elements_by_tag_name('tr')
        # Get all the links
        links = get_column_link_from_entries(table, rows)
        # Get Risk reward and potensial
        columns = get_columns_from_each_entry_in_list(table, rows)
        # Loop through each links
        for count, link in enumerate(links):
            # Go to kort sikt
            driver.get(link)
            # Pause for a moment
            time.sleep(3)
            # Get the data
            data = get_data_from_kortsikt(count, columns)
            # Add the data to the llist
            dataList.append(data)
        # Increase by the offset valeu
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
    driver.implicitly_wait(10)
    
    # Login
    login()
    
    # Sleep for a moment before the script continues
    time.sleep(5)

    # Get the data returned from handelsmuligheter
    data = handelsmuligheter()

    # Export the data to excel
    export_to_excel(data)

except NoSuchElementException as e:
    print("error: ", e)
