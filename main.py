import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException

markets = {
    'Oslo Bors': '//*[@id="market_dropdown"]/a[1]',
    'Stockholmsborsen': '//*[@id="market_dropdown"]/a[2]',
    'Kobenhavn Fondsbors': '//*[@id="market_dropdown"]/a[3]',
    'Investeringsforeninger': '//*[@id="market_dropdown"]/a[4]',
    'Helsingin pörssi': '//*[@id="market_dropdown"]/a[5]',
    'World Indices': '//*[@id="market_dropdown"]/a[6]',
    'US Stocks': '//*[@id="market_dropdown"]/a[7]',
    'US 30': '//*[@id="market_dropdown"]/a[8]',
    'Nasdaq 100': '//*[@id="market_dropdown"]/a[9]',
    'US 500': '//*[@id="market_dropdown"]/a[10]',
    'Toronto Stock Exchange': '//*[@id="market_dropdown"]/a[11]',
    'London Stock Exchange': '//*[@id="market_dropdown"]/a[12]',
    'Euronext Amsterdam': '//*[@id="market_dropdown"]/a[13]',
    'Euronext Brussel': '//*[@id="market_dropdown"]/a[14]',
    'DAX 30': '//*[@id="market_dropdown"]/a[15]',
    'TECDAX': '//*[@id="market_dropdown"]/a[16]',
    'Frankfurt': '//*[@id="market_dropdown"]/a[17]',
    'MDAX': '//*[@id="market_dropdown"]/a[18]',
    'CDAX': '//*[@id="market_dropdown"]/a[19]',
    'SDAX': '//*[@id="market_dropdown"]/a[20]',
    'Prime Standard': '//*[@id="market_dropdown"]/a[21]',
    'CAC 40': '//*[@id="market_dropdown"]/a[22]',
    'Mumbai S.E.': '//*[@id="market_dropdown"]/a[23]',
    'National S.E.': '//*[@id="market_dropdown"]/a[24]',
    'Commodities': '//*[@id="market_dropdown"]/a[25]',
    'Currency': '//*[@id="market_dropdown"]/a[26]',
    'Cryptopcurrency': '//*[@id="market_dropdown"]/a[27]',
}

def stock_pick_page(driver, key):


    def get_data_from_graferkortsikt():
        
        # Initialize an empty list
        data_list = []
        # Initialize an empty dictionary
        data_dict = {}
        # Get Kurs
        sistekurs = driver.find_element_by_xpath('//*[@id="ca2017_HeadPriceAndDateInfo"]')
        kurs = sistekurs.text.split()[2]
        # Get name
        name = driver.find_element_by_xpath('//*[@id="ca2017_HeadNameTicker"]')
        # Add name to dictionary
        data_dict['name'] = name.text
        # Add kurs to dictionary
        data_dict['kurs'] = kurs

        # Loop through table
        table = driver.find_element_by_xpath('//*[@id="ca2017_QuantIndicatorsTable"]')
        for row in table.find_elements_by_css_selector('tr'):
            for cell in row.find_elements_by_tag_name('td'):
                if 'Trendgulv' in cell.text:
                    # Get only the values from text (trendgulv, trendtak)
                    trendgulv = cell.text.split('\n')[1]
                    trendtak = cell.text.split('\n')[2]
                    # Add to dictionary
                    data_dict['trendgulv'] = trendgulv 
                    data_dict['trendtak'] = trendtak 
                    # Add dictionary to list
                    data_list.append(data_dict)
                elif 'volumbalanse' in cell.text:
                    # Get only the value from text (volumbalanse)
                    volum =  cell.text.split()[-1]
                    # Add volumne to dictionary
                    data_dict['volum'] = volum 
                    # Add dictionary to list
                    data_list.append(data_dict)
                elif 'momentum' in cell.text:
                    # Get only the value from text (KSI)
                    ksi =  cell.text.split()[-1]
                    # Add KSI to dictionary
                    data_dict['ksi'] = ksi
                    # Add dictionary to list
                    data_list.append(data_dict)

        return data_list

    def display_table_graferkortsikt(select, option):
        # Click on select
        elem = driver.find_element_by_xpath(select)
        elem.click()

        # Click on option
        elem = driver.find_element_by_xpath(option)
        elem.click()

        # Click on kort sikt
        elem = driver.find_element_by_xpath('//*[@id="productBox"]/table[3]/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[1]/td/a')
        elem.click()

        get_data_from_graferkortsikt()

    if key == 1:
        display_table_graferkortsikt('//*[@id="tradingOpportunitiesAdvancedOptions"]/tbody/tr[1]/td[3]/form[2]/select',
        '//*[@id="tradingOpportunitiesAdvancedOptions"]/tbody/tr[1]/td[3]/form[2]/select/option[4]')
    
    elif key == 2:
        display_table_graferkortsikt('//*[@id="signalAdvancedOptions"]/tbody/tr[1]/td[4]/form[2]/select', 
        '//*[@id="signalAdvancedOptions"]/tbody/tr[1]/td[4]/form[2]/select/option[4]')

    elif key == 3:
        display_table_graferkortsikt('//*[@id="trendsignalAdvancedOptions"]/tbody/tr[1]/td[4]/form/select', 
        '//*[@id="trendsignalAdvancedOptions"]/tbody/tr[1]/td[4]/form/select/option[4]')

    elif key == 4:
        display_table_graferkortsikt('//*[@id="pivotPointsAdvancedOptions"]/tbody/tr[1]/td[5]/form/select', 
        '//*[@id="pivotPointsAdvancedOptions"]/tbody/tr[1]/td[5]/form/select/option[4]')
          

try: 
    # Launch
    browserPath = 'C:\Program Files (x86)\chromedriver.exe'
    driver = webdriver.Chrome(browserPath)
    driver.get('https://investtech.com/no/market.php?MarketID=1&product=17')
    driver.implicitly_wait(10)

    elem = driver.find_element_by_xpath('//*[@id="login_button"]')

    elem.click()

    # Login
    username = driver.find_element_by_xpath('//*[@id="login_form"]/div[1]/input')
    password = driver.find_element_by_xpath('//*[@id="login_form"]/div[2]/input')

    username.send_keys('90079450')
    password.send_keys('vinter1234')

    elem = driver.find_element_by_xpath('//*[@id="login_form_submit"]')
    elem.click()

    
    # Navigate

    for key, value in markets.items():

        time.sleep(3)

        # Click on dropdown button
        elem = driver.find_element_by_class_name('dropbtn')
        elem.click()

        time.sleep(3)

        # Click on dropdown button
        elem = driver.find_element_by_xpath(value)
        elem.click()
        
        time.sleep(3)

        # Click on stockpicking
        elem = driver.find_element_by_xpath('//*[@id="navigationmenu"]/li[4]/div/button/a')
        elem.click()

        time.sleep(3)

        menu = driver.find_element_by_xpath('//*[@id="navigationsubmenu"]')

        # Handelsmuligheter
        handelsmuligheter = driver.find_element_by_link_text('Handelsmuligheter')
        handelsmuligheter.click()
        
        stock_pick_page(driver, 1)

        # Signaler
        signaler = driver.find_element_by_link_text('Signaler')
        signaler.click()

        stock_pick_page(driver, 2)

        # Trendsignaler
        trendsignaler = driver.find_element_by_link_text('Trendsignaler')
        trendsignaler.click()

        stock_pick_page(driver, 3)

        # Pivotpunkter
        pivotpunkter = driver.find_element_by_link_text('Pivotpunkter')
        pivotpunkter.click()
        
        stock_pick_page(driver, 4)
        
        

except NoSuchElementException as e:
    print("error: ", e)
