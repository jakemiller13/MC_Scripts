# -*- coding: utf-8 -*-
"""
Created on Tue Dec 18 11:19:57 2018

@author: jmiller
"""

# Import packages
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.common.exceptions import NoSuchWindowException,\
                                       NoSuchElementException,\
                                       WebDriverException,\
                                       NoAlertPresentException
import datetime
import time
import traceback
import pandas as pd

# Folder location of saved data
folder_location = 'S:\Engineering\Jake\mc-missing-data'

# Define driver and waits
driver = webdriver.Chrome(executable_path =\
                          r'S:\Engineering\Jake\mc-missing-data\chromedriver')
wait = WebDriverWait(driver, 5)

class MissingInfoException(Exception):
    '''
    Raised if an alert is raised while trying to save accession form
    '''
    def __init__(self, message):
        self.message = message

class NoCheckedOutFormException(Exception):
    '''
    Raised if there is no form to check back in
    '''
    def __init__(self, message):
        self.message = message

def load_data(folder_location):
    '''
    Returns dataframe of corrected data
    File name format: "Corrected_Data_YYYY_MM_DD.csv"
    NOTE: The date used will be TODAYS date. This is to ensure most recent
          information is used
    '''
    folder = folder_location
    file_name = 'Corrected_Data_' + \
            str(datetime.datetime.now().strftime('%Y_%m_%d')) + '.csv'
    data = pd.read_csv(folder + '\\' + file_name, na_values = (''), 
                       keep_default_na = False)
    return data.drop(data.columns[0], axis = 1)

def get_user_pw():
    '''
    Asks for user_id and password
    Returns user_id, password
    '''
    user_id = input('User ID: ')
    password = input('Password: ')
    return user_id, password

def load_mastercontrol(driver, user_id, password, dev = True):
    '''
    Opens MasterControl using "user_id", "password"
    Default is development environment. Specify dev = False to load production
    '''
    if dev:
        url = 'https://acceleronpharma.mastercontrol.com/mcdev/login/'
    else:
        url = 'https://acceleronpharma.mastercontrol.com/mc/login/'
    driver.get(url)
    driver.find_element_by_name('userid').send_keys(user_id)
    driver.find_element_by_name('password').send_keys(password)
    driver.find_element_by_id('MCP_BUTTON_40').click()

def switch_to_window(driver, window):
    '''
    Switches between windows
    Used because sometimes accession form opens in new window, sometimes not
    '''
    if window == 'Last':
        driver.switch_to_window(driver.window_handles[-1])
    elif window == 'First':
        driver.switch_to_window(driver.window_handles[0])

def switch_to_frame(driver, frame = 'None'):
    '''
    Navigates to various frames within MasterControl
    Default is "None" which switches to default_content
    '''
    switch_to_window(driver, 'Last')
    driver.switch_to_default_content()
    if frame == 'myframe':
        driver.switch_to_frame('myframe')
    if frame == 'bigFrame':
        driver.switch_to_frame('bigFrame')

def asset_search(driver, asset_number):
    '''
    Searches for "asset_number"
    '''
    switch_to_frame(driver)
    driver.find_element_by_name('strSearch').clear()
    driver.find_element_by_name('strSearch').send_keys(asset_number)
    driver.find_element_by_xpath('//input[@value = "Go"]').click()

def open_forms_folder(driver):
    '''
    If >25 forms, expands search
    Otherwise, opens "Forms" folder
    '''
    switch_to_frame(driver, 'myframe')
    try:
        driver.find_element_by_xpath(
                '//div[@id = "folder_Forms"]/'\
                'a[contains(text(), "Expand this search")]').click()
        return
    except NoSuchElementException:
        try:
            driver.execute_script("toggleFolder('image_Forms','group_Forms');")
        except WebDriverException:
            return

def get_total_pages(driver):
    '''
    Returns total number of pages, if greater than 1
    '''
    switch_to_frame(driver, 'myframe')
    return len(Select(driver.find_element_by_name('page')).options)

def select_page(driver, page_num):
    '''
    Switches to selected page, if available
    Returns "break" if no more pages to select
    '''
    switch_to_frame(driver, 'myframe')
    Select(driver.find_element_by_name('page')).select_by_value(str(page_num))

def open_infocard(driver, asset_number):
    '''
    Opens InfoCard for "asset_number"
    '''
    switch_to_frame(driver, 'myframe')
    driver.find_element_by_link_text(asset_number).click()

def checkout_form(driver, asset_number):
    '''
    Checks out and opens asset accession form
    If form is already checked out, attempts to check in before checking out
    '''
    switch_to_frame(driver, 'myframe')
    driver.find_element_by_id('revision').click()
    try:
        driver.find_element_by_link_text('Check Out').click()
        driver.find_element_by_id('button_checkout').click()
    except NoSuchElementException:
        checkin_form(driver, asset_number)
        checkout_form(driver)

def switch_to_page(driver, page):
    '''
    Switches between three accession form pages
    '''
    if page == 1:
        driver.find_element_by_id('formTabsnav1').click()
    if page == 2:
        driver.find_element_by_id('formTabsnav2').click()
    if page == 3:
        driver.find_element_by_id('formTabsnav3').click()

def save_accession_form(driver, asset_number):
    '''
    Saves accession form
    If any pop-up happens here:
        -acknowledges pop-up
        -raises and records exception
        -closes window WITHOUT saving
    '''
    driver.find_element_by_id('button_save').click()
    try:
        WebDriverWait(driver, 1).until(EC.alert_is_present())
        checkin_form(driver)
        raise MissingInfoException(str(asset_number) +\
                                   ': Missing information, not saved')
    except (NoAlertPresentException, NoSuchWindowException):
        switch_to_frame(driver, 'myframe')
        driver.find_element_by_id('button_save').click()
        driver.find_element_by_id('approve').click()
        driver.find_element_by_link_text('Quick Approve').click()

def approve_form(driver, user_id, password):
    '''
    Insert comments "Updated all missing information"
    Input user_id, password
    '''
    switch_to_frame(driver, 'bigFrame')
    driver.find_element_by_id('comments').\
                              send_keys('Updated all missing information')

    try:
        driver.find_element_by_name('userID').send_keys(user_id)
    except NoSuchElementException:
        pass
    driver.find_element_by_name('esig').send_keys(password)
    driver.find_element_by_name('btSubmit').click()
    
def checkin_form(driver, asset_number):
    '''
    Checks in form
    Called in instances where cannot check out form
    '''
    try:
        driver.switch_to_alert().accept
        driver.close()
        switch_to_frame(driver, 'myframe')
        driver.find_element_by_id('cancel').click()
    except (NoSuchElementException, NoAlertPresentException):
        pass
    finally:
        switch_to_frame(driver, 'myframe')
        driver.find_element_by_id('revision').click()
        driver.find_element_by_link_text('Check In').click()
        driver.find_element_by_id('check_in_wo').click()
        driver.find_element_by_id('button_checkin').click()
        driver.switch_to_default_content()
    if driver.find_element_by_id(\
              'messageDisplayed').get_attribute('value') == 'true':
        raise NoCheckedOutFormException(str(asset_number) +\
                                   ': Has no form to check in')

def assign_data_fields(df, asset_number):
    '''
    Assigns data from df to respective field
    '''
    data_dict = df.loc[df['asset_number'] == int(asset_number)].\
                       to_dict('records')
    return data_dict[0]

def data_per_page():
    '''
    Returns 3 lists, 1 for each page in accession form
    Easier to navigate and upload data per page
    '''
    page_1 = ['asset_number',
              'title',
              'department',
              'manufacturer',
              'manufacturer_model_number',
              'manufacturer_serial_number',
              'vendor',
              'vendor_catalog_number',
              'purchase_order_number',
              'portable',
              'asset_description',
              'intended_use',
              'process_range',
              'classification',
              'new_SOP',
              'new_logbooks']
    page_2 = ['calibration_required',
              'calibration_justification_no',
              'calibration_description',
              'calibration_frequency',
              'maintenance_required',
              'maintenance_justification_no',
              'maintenance_description',
              'maintenance_frequency',
              'scan_attached',
              'tier_level',
              'qual_val_needed',
              'validation_justification_no',
              'validation_description']
    page_3 = ['protocol_number',
              'comments']
    return page_1, page_2, page_3

def complete_page_1(asset_info, page):
    '''
    General Information tab
    Cannot create loop because data in different forms (text, radio, etc)
    '''
    if driver.find_element_by_id('mastercontrol.form.number').\
                                 get_attribute('value') == '':
        driver.find_element_by_id('mastercontrol.form.number').\
                                  send_keys(asset_info[page[0]])
                                  
    if driver.find_element_by_id('txtTitle').\
                                 get_attribute('value') == '':
        driver.find_element_by_id('txtTitle').\
                                  send_keys(asset_info[page[1]])
                                  
    if Select(driver.find_element_by_id(\
              'mastercontrol.dataset.recordids.Department')).\
              first_selected_option.text == '':
        Select(driver.find_element_by_id(\
               'mastercontrol.dataset.recordids.Department')).\
               select_by_value(asset_info[page[2]])
               
    if driver.find_element_by_id('txtManufacturer').\
                                 get_attribute('value') == '':
        driver.find_element_by_id('txtManufacturer').\
                                  send_keys(asset_info[page[3]])
                                  
    if driver.find_element_by_id('txtManufacturerModelNumber').\
                                 get_attribute('value') == '':
        driver.find_element_by_id('txtManufacturerModelNumber').\
                                  send_keys(asset_info[page[4]])
                                  
    if driver.find_element_by_id('txtManufacturerSerialNumber').\
                                 get_attribute('value') == '':
        driver.find_element_by_id('txtManufacturerSerialNumber').\
                                  send_keys(asset_info[page[5]])
                                  
    if driver.find_element_by_id('txtVendor').\
                                 get_attribute('value') == '':
        driver.find_element_by_id('txtVendor').\
                                  send_keys(asset_info[page[6]])
                                  
    if driver.find_element_by_id('txtVendorCatalogNumber').\
                                 get_attribute('value') == '':
        driver.find_element_by_id('txtVendorCatalogNumber').\
                                  send_keys(asset_info[page[7]])

    if driver.find_element_by_id('txtPurchaseOrderNumber').\
                                 get_attribute('value') == '':
        driver.find_element_by_id('txtPurchaseOrderNumber').\
                                  send_keys(asset_info[page[8]])

    try:
        driver.find_element_by_xpath(\
               '//input[@id = "rbPortable" and @checked = "true"]')
    except NoSuchElementException:
        driver.find_element_by_xpath(
                '//input[@id = "rbPortable" and @value = "' +\
                asset_info[page[9]] + '"]').click()
        
    if driver.find_element_by_id('txtAssetDescription').\
                                 get_attribute('value') == '':
        driver.find_element_by_id('txtAssetDescription').\
                                  send_keys(asset_info[page[10]])
                                  
    if driver.find_element_by_id('txtIntendedUse').\
                                 get_attribute('value') == '':
        driver.find_element_by_id('txtIntendedUse').\
                                  send_keys(asset_info[page[11]])
                                  
    if driver.find_element_by_id('txtProcessRange').\
                                 get_attribute('value') == '':
        driver.find_element_by_id('txtProcessRange').\
                                  send_keys(asset_info[page[12]])
                                  
    try:
        driver.find_element_by_xpath(\
               '//input[@id = "rbClassification" and @checked = "true"]')
    except NoSuchElementException:
        driver.find_element_by_xpath(
                '//input[@id = "rbClassification" and @value = "' +\
                asset_info[page[13]] + '"]').click()
        
    try:
        driver.find_element_by_xpath(\
               '//input[@id = "rbYes/No3" and @checked = "true"]')
    except NoSuchElementException:
        driver.find_element_by_xpath(
                '//input[@id = "rbYes/No3" and @value = "' +\
                asset_info[page[14]] + '"]').click()
        
    try:
        driver.find_element_by_xpath(\
               '//input[@id = "rbYes/No4" and @checked = "true"]')
    except NoSuchElementException:
        driver.find_element_by_xpath(
                '//input[@id = "rbYes/No4" and @value = "' +\
                asset_info[page[15]] + '"]').click()

def complete_page_2(asset_info, page):
    '''
    Calibration & Maintenance tab
    Cannot create loop because data in different forms (text, radio, etc)
    '''
    try:
        driver.find_element_by_xpath(\
               '//input[@id = "rbAssetRequiresScheduledCalibration" and '\
               '@checked = "true"]')
    except NoSuchElementException:
        driver.find_element_by_xpath(\
                '//input[@id = "rbAssetRequiresScheduledCalibration" and '\
                '@value = "' +\
                asset_info[page[0]] + '"]').click()
    
    if driver.find_element_by_id('txtJustificationNoCal').\
                                 get_attribute('value') == '':
        driver.find_element_by_id('txtJustificationNoCal').\
                                  send_keys(asset_info[page[1]])

    if driver.find_element_by_id('txtDescriptionofCal').\
                                 get_attribute('value') == '':
        driver.find_element_by_id('txtDescriptionofCal').\
                                  send_keys(asset_info[page[2]])

    if driver.find_element_by_id('txtFrequencyofCal').\
                                 get_attribute('value') == '':
        driver.find_element_by_id('txtFrequencyofCal').\
                                  send_keys(asset_info[page[3]])

    try:
        driver.find_element_by_xpath(\
               '//input[@id = "rbAssetRequiresScheduledMaintenance" and '\
               '@checked = "true"]')
    except NoSuchElementException:
        driver.find_element_by_xpath(\
                '//input[@id = "rbAssetRequiresScheduledMaintenance" and '\
                '@value = "' +\
                asset_info[page[4]] + '"]').click()

    if driver.find_element_by_id('txtJustificationNoMaintenance').\
                                 get_attribute('value') == '':
        driver.find_element_by_id('txtJustificationNoMaintenance').\
                                  send_keys(asset_info[page[5]])

    if driver.find_element_by_id('txtDescriptionofMaintenance').\
                                 get_attribute('value') == '':
        driver.find_element_by_id('txtDescriptionofMaintenance').\
                                  send_keys(asset_info[page[6]])

    if driver.find_element_by_id('txtFrequencyofScheduledMaintenance').\
                                 get_attribute('value') == '':
        driver.find_element_by_id('txtFrequencyofScheduledMaintenance').\
                                  send_keys(asset_info[page[7]])

    try:
        driver.find_element_by_xpath(\
               '//input[@id = "rbInitialCalMaintenance" and '\
               '@checked = "true"]')
    except NoSuchElementException:
        driver.find_element_by_xpath(\
                '//input[@id = "rbInitialCalMaintenance" and '\
                '@value = "' +\
                asset_info[page[8]] + '"]').click()

    try:
        driver.find_element_by_xpath(\
               '//input[@id = "rbTierLevel" and '\
               '@checked = "true"]')
    except NoSuchElementException:
        driver.find_element_by_xpath(\
                '//input[@id = "rbTierLevel" and '\
                '@value = "' +\
                asset_info[page[9]] + '"]').click()

    try:
        driver.find_element_by_xpath(\
               '//input[@id = "rbQualVal Needed" and '\
               '@checked = "true"]')
    except NoSuchElementException:
        driver.find_element_by_xpath(\
                '//input[@id = "rbQualVal Needed" and '\
                '@value = "' +\
                asset_info[page[10]] + '"]').click()

    if driver.find_element_by_id('txtJustificationNoVal').\
                                 get_attribute('value') == '':
        driver.find_element_by_id('txtJustificationNoVal').\
                                  send_keys(asset_info[page[11]])

    if driver.find_element_by_id('txtBRIEFDescriptionofQual/Val').\
                                 get_attribute('value') == '':
        driver.find_element_by_id('txtBRIEFDescriptionofQual/Val').\
                                  send_keys(asset_info[page[12]])

def complete_page_3(asset_info, page):
    '''
    Qualification/Validation tab
    Cannot create loop because data in different forms (text, radio, etc)
    '''
    if driver.find_element_by_id('txtProtocolNumber').\
                                 get_attribute('value') == '':
        driver.find_element_by_id('txtProtocolNumber').\
                                  send_keys(asset_info[page[0]])

    if driver.find_element_by_id('txtComments').\
                                 get_attribute('value') == '':
        driver.find_element_by_id('txtComments').\
                                  send_keys(asset_info[page[1]])

def execute_program(driver, folder_location, assets):
    '''
    Program execution
    Separate function so we can time it
    '''
    logf = open("error_log.txt", "w")
    log_correct = open("completed_assets.txt", "w")
    
    df = load_data(folder_location)
    page_1, page_2, page_3 = data_per_page()
    user_id, password = get_user_pw()
    load_mastercontrol(driver, user_id, password, dev = False)
    
    for asset_number in range(1816, assets):
        asset_number = str(asset_number).rjust(4, '0')
        try:
            try:
                asset_info = assign_data_fields(df, asset_number)
            except IndexError:
                continue
            asset_search(driver, asset_number)
            open_forms_folder(driver)
            
            try:
                total_pages = get_total_pages(driver)
            except NoSuchElementException:
                total_pages = 1
            
            page_num = 1
            
            while page_num <= total_pages:
                try:
                    select_page(driver, page_num)
                except NoSuchElementException:
                    pass
                try:
                    open_infocard(driver, asset_number)
                except NoSuchElementException:
                    page_num += 1
            
            checkout_form(driver, asset_number)
            switch_to_window(driver, 'Last')
            complete_page_1(asset_info, page_1)
            switch_to_page(driver, 2)
            complete_page_2(asset_info, page_2)
            switch_to_page(driver, 3)
            complete_page_3(asset_info, page_3)
            save_accession_form(driver, asset_number)
            approve_form(driver, user_id, password)
            print(('{0}: Information Updated Successfully\n'.\
                              format(str(asset_number))))
            log_correct.write('{0}: Information Updated Successfully\n'.\
                              format(str(asset_number)))
        except Exception as e:
            print('{0}: Error Updating - See Error Log\n'.\
                              format(str(asset_number)))
            log_correct.write('{0}: Error Updating - See Error Log\n'.\
                              format(str(asset_number)))
            logf.write('{0}:\n{1} {2}\n'.format(str(asset_number),
                   str(traceback.format_exc()), str(e)))

# Start time for program execution
start_time = time.time()
print('Start time: ' + time.asctime())

execute_program(driver, folder_location, 1900)

# End time for program execution
end_time = time.time()
print('End time: ' + time.asctime())

# Print out total execution time
total_time = end_time - start_time
hours = total_time // 3600
minutes = (total_time - hours * 3600) // 60
seconds = (total_time - hours * 3600 - minutes * 60)

print('Execution time: ' + str(int(hours)) + 'h:' +
      f'{int(minutes):02}' + 'm:' + f'{int(seconds):02}s')