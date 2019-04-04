# -*- coding: utf-8 -*-
"""
Created on Tue Dec  4 10:31:16 2018

@author: jmiller
"""

# Import packages
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.common.exceptions import UnexpectedAlertPresentException,\
                                       WebDriverException,\
                                       NoSuchElementException
from selenium.webdriver.common.alert import Alert
import time
import pandas as pd

# Define program parameters
driver = webdriver.Chrome(executable_path =
                          r'S:\Engineering\Jake\mc-correcting-asset-locations\\'
                          'chromedriver')
wait = WebDriverWait(driver, 5)
file_path = 'tier 1 and 2 room locations.csv'

# Are you working in PRODUCTION or DEVELOPMENT environment?
url = 'https://acceleronpharma.mastercontrol.com/mc/login/'
#url = 'https://acceleronpharma.mastercontrol.com/mcdev/login/'

def load_data(file_path):
    '''
    Loads in and cleans up data
    '''
    df = pd.read_csv(file_path, na_values = (''), keep_default_na = False,
                     dtype = {'Asset Number':str})
    df = df[['Asset Number', 'Room Number']]
    df.dropna(inplace = True)
    return df

def load_mastercontrol(driver, url):
    '''
    Opens MasterControl
    Asks for user_id and password
    Times out after 30 seconds
    '''
    user_id = input('User Name: ')
    password = input('Password: ')
    driver.get(url)
    driver.find_element_by_name('userid').send_keys(user_id)
    driver.find_element_by_name('password').send_keys(password)
    driver.find_element_by_id('MCP_BUTTON_40').click()
    WebDriverWait(driver, 30).until(EC.url_changes(url))
    print('\n' * 1000) #NOTE: there is definitely a better way to do this

def asset_search(driver, asset_number):
    '''
    Searches for "asset_number"
    '''
    driver.switch_to_default_content()
    driver.find_element_by_name('strSearch').clear()
    driver.find_element_by_name('strSearch').send_keys(asset_number)
    driver.find_element_by_xpath('//input[@value = "Go"]').click()

def open_forms_folder(driver):
    '''
    Opens "Forms" folder
    '''
    driver.switch_to_frame('myframe')
    try:
        driver.execute_script("toggleFolder('image_Forms','group_Forms');")
    except WebDriverException:
        return

def open_infocard(driver, asset_number):
    '''
    Opens InfoCard for "asset_number"
    '''
    driver.find_element_by_link_text(asset_number).click()
    
def edit_infocard(driver):
    '''
    Begins editing infocard
    '''
    driver.find_elements_by_link_text('Edit')[0].click()
    driver.find_elements_by_link_text('Edit')[1].click()

def nav_custom_fields(driver):
    '''
    Navigates to Custom Fields tab
    '''
    driver.find_element_by_id('customFields').click()
    
def update_functional_location(driver, functional_location):
    '''
    Updates "Generator" and "Electrical Panel" fields
    '''
    driver.switch_to_default_content()
    driver.switch_to_frame('myframe')
    driver.switch_to_frame('addInfo')
    for option in Select(driver.find_element_by_id(\
                         'cfName_Functional Location')).options:
        if functional_location in option.text:
            Select(driver.find_element_by_id('cfName_Functional Location')).\
                   select_by_visible_text(option.text)

def clear_location(driver):
    '''
    Clears old location information
    '''
    driver.find_element_by_id('cfName_Location').clear()

def save(driver):
    '''
    Saves InfoCard
    '''
    driver.switch_to_default_content()
    driver.switch_to_frame('myframe')
    driver.find_element_by_id('save').click()
    driver.switch_to_default_content()
    driver.switch_to_frame('auditFrame')
    driver.find_element_by_name('Chg_reason').\
           send_keys('Updated functional location')
    driver.find_element_by_name('ok').click()

def program(driver, file_path, url):
    '''
    Main program
    '''
    logf = open("error_log.txt", "w")
    count = 0
    
    df = load_data(file_path)
    load_mastercontrol(driver, url)
    
    for asset in df['Asset Number']:
        try:        
            functional_location = df.loc[df['Asset Number'] == asset,\
                                            'Room Number'].item()
                
            asset_number = asset.zfill(4)
            asset_search(driver, asset_number)
            open_forms_folder(driver)
            open_infocard(driver, asset_number)
            try:
                edit_infocard(driver)
                nav_custom_fields(driver)
            except UnexpectedAlertPresentException:
                Alert(driver).accept()
                print('Unable to update || Asset Number\
                      [' + asset_number + ']')
                continue
            update_functional_location(driver, functional_location)
            try:
                clear_location(driver)
            except NoSuchElementException:
                pass
            save(driver)
            
            count += 1
            
            print('Successfully updated || Asset Number [' + asset_number + ']\
                  || Functional Location [' + functional_location +']')
            
        except Exception as e:
            logf.write(asset + ':\n' + str(e))
    
    print('Corrected ' + str(count) + ' InfoCards')
    logf.close()

def run(program):
    '''
    Runs program and measures how long it takes
    '''
    start_time = time.time()
    print('Start time: ' + time.asctime())
    
    program(driver, file_path, url)
    
    end_time = time.time()
    print('End time: ' + time.asctime())
    
    total_time = end_time - start_time
    hours = total_time // 3600
    minutes = (total_time - hours * 3600) // 60
    seconds = (total_time - hours * 3600 - minutes * 60)
    print('Execution time: ' + str(int(hours)) + 'h:' + f'{int(minutes):02}' + 'm:' + f'{int(seconds):02}s')

#run(program)