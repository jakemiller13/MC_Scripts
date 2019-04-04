# -*- coding: utf-8 -*-
"""
Created on Tue Oct 23 13:31:22 2018

@author: jmiller
"""

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchWindowException,\
                                       NoSuchFrameException,\
                                       NoSuchElementException,\
                                       WebDriverException
import csv
import traceback
import time
import datetime

combined_asset_info = [['asset_number',
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
                        'new_logbooks',
                        'calibration_required',
                        'calibration_justification_no',
                        'calibration_description',
                        'calibration_frequency',
                        'maintenance_required',
                        'maintenance_justification_no',
                        'maintenance_description',
                        'maintenance_frequency',
                        'scan_attached',
                        'scan_file',
                        'tier_level',
                        'qual_val_needed',
                        'validation_justification_no',
                        'validation_description',
                        'protocol_number',
                        'comments']]

# Define driver and waits
driver = webdriver.Chrome(executable_path = 
                          r'S:\Engineering\Jake\mc-missing-data\chromedriver')
wait = WebDriverWait(driver, 5)

# Class to skip asset number if it can't be found
class NoAsset(Exception):
    pass

# Navigational functions
def load_mastercontrol(driver):
    '''
    Opens MasterControl, waits up to 30 seconds for username/password entry
    '''

    url = 'https://acceleronpharma.mastercontrol.com/mc/login/'
    driver.get(url)
    WebDriverWait(driver, 30).until(EC.url_changes(url))

def close_extra_windows():
    '''
    Closes any extra windows that may have opened before starting next loop
    '''
    while len(driver.window_handles) > 1:
        driver.switch_to_window(driver.window_handles[-1])
        driver.close()

def navigate_to_frame(frame = None):
    '''
    Navigates to "frame"
    Default is "default_content" i.e. main window
    '''
    driver.switch_to_default_content()
    if frame == 'myframe' or frame == 'formFrame':
        driver.switch_to_frame('myframe')
        if frame == 'formFrame':
            driver.switch_to_frame('formFrame')

def asset_search(asset_number):
    '''
    Searches for asset number
    '''
    driver.switch_to_window(driver.window_handles[0])
    navigate_to_frame()
    driver.find_element_by_name('strSearch').clear()
    driver.find_element_by_name('strSearch').send_keys(asset_number)
    driver.find_element_by_xpath('//input[@value = "Go"]').click()

def open_forms_folder():
    '''
    Opens "Forms" folder if it exists
    '''
    navigate_to_frame('myframe')
    try:
        driver.execute_script("toggleFolder('image_Forms','group_Forms');")
    except WebDriverException:
        return

def expand_search():
    '''
    If "Forms" folder has more than 25 records, expands folder
    '''
    navigate_to_frame('myframe')
    try:
        driver.find_element_by_xpath("//div[@id = 'folder_Forms']/a").click()
    except NoSuchElementException:
        return

def locate_asset(asset_number):
    '''
    Checks current page for "asset_number" and opens if located
    If not located, advances page up to "total_pages"
    '''
    navigate_to_frame('myframe')
    page_number = 1
    try:
        total_pages = len(Select(driver.find_element_by_name('page')).options)
    except NoSuchElementException:
        total_pages = 1
    try:
        select_page(page_number)
    except NoSuchElementException:
        pass        
    while page_number <= total_pages:
        try:
            open_asset_accession_form(asset_number)
            return
        except NoSuchElementException:
            try:
                page_number += 1
                select_page(page_number)
            except NoSuchElementException:
                pass
    raise NoAsset

def open_asset_accession_form(asset_number):
    '''
    Opens accession form if located
    '''
    navigate_to_frame('myframe')
    driver.find_element_by_link_text(asset_number).click()
    driver.find_element_by_id('MCD_ALT_TEXT_30').click()

def select_page(page_number):
    '''
    Selects page number while trying to locate asset
    '''
    navigate_to_frame('myframe')
    Select(driver.find_element_by_name(
            'page')).select_by_value(str(page_number))

def navigate_to_accession_form():
    '''
    Navigates driver to accession form
    Keeps checking if Title has shown up until it does
    Necessary because forms randomly open in same window, sometimes new window
    '''
    while True:
        try:
            driver.switch_to_window(driver.window_handles[-1])
            try:
                navigate_to_frame('formFrame')
            except NoSuchFrameException:
                pass
            try:
                driver.find_element_by_id('txtTitle')
                break
            except NoSuchElementException:
                pass
#            if len(driver.window_handles) == 1:
#                navigate_to_frame('formFrame')
#            driver.find_element_by_id('txtTitle')
#            break
        except (NoSuchWindowException, NoSuchFrameException,
                NoSuchElementException):
            pass

def gather_data():
    '''
    Gets all info from accession form
    NOTE: this function needs to be cleaned up, e.g.:
        -radio buttons
        -length of text in lines
    '''
    asset_info = []
    
    # General Information tab
    title = driver.find_element_by_id('txtTitle').get_attribute('value')
    
    # Get department info - "selected" doesn't exist if not selected, so need try/except
    try:
        department = driver.find_element_by_xpath(
                "//select[@id = 'mastercontrol.dataset.recordids.Department']\
                /option[@selected = 'true']").text
    except NoSuchElementException:
        department = ''
    
    manufacturer = driver.find_element_by_id(
            'txtManufacturer').get_attribute('value')
    manufacturer_model_number = driver.find_element_by_id(
            'txtManufacturerModelNumber').get_attribute('value')
    manufacturer_serial_number = driver.find_element_by_id(
            'txtManufacturerSerialNumber').get_attribute('value')
    vendor = driver.find_element_by_id(
            'txtVendor').get_attribute('value')
    vendor_catalog_number = driver.find_element_by_id(
            'txtVendorCatalogNumber').get_attribute('value')
    purchase_order_number = driver.find_element_by_id(
            'txtPurchaseOrderNumber').get_attribute('value')

    # Portable - radio button - need to determine if "Yes" or "No"
    portable_list = driver.find_elements_by_id('rbPortable')
    for i in portable_list:
        if i.get_attribute('checked') == 'true':
            portable = i.get_attribute('value')
            break
        else:
            portable = ''
    
    asset_description = driver.find_element_by_id(
            'txtAssetDescription').get_attribute('value')
    intended_use = driver.find_element_by_id(
            'txtIntendedUse').get_attribute('value')
    process_range = driver.find_element_by_id(
            'txtProcessRange').get_attribute('value')

    # Classification - radio button - need to determine if "Yes" or "No"
    classification_list = driver.find_elements_by_id('rbClassification')
    for i in classification_list:
        if i.get_attribute('checked') == 'true':
            classification = i.get_attribute('value')
            break
        else:
            classification = ''
    
    # New SOP - radio button - need to determine if "Yes" or "No"
    new_SOP_list = driver.find_elements_by_id('rbYes/No3')
    for i in new_SOP_list:
        if i .get_attribute('checked') == 'true':
            new_SOP = i.get_attribute('value')
            break
        else:
            new_SOP = ''
    
    # New logbook - adio button - need to determine if "Yes" or "No"
    new_logbooks_list = driver.find_elements_by_id('rbYes/No4')
    for i in new_logbooks_list:
        if i.get_attribute('checked') == 'true':
            new_logbooks = i.get_attribute('value')
            break
        else:
            new_logbooks = ''
    
    # Switch to Calibration & Maintenance tab
    calibration_maintenance_tab = driver.find_element_by_id('formTabsnav2')
    calibration_maintenance_tab.click()
    
    # Calibration required - radio button - need to determine if "Yes" or "No"
    calibration_required_list = driver.find_elements_by_id(
            'rbAssetRequiresScheduledCalibration')
    for i in calibration_required_list:
        if i.get_attribute('checked') == 'true':
            calibration_required = i.get_attribute('value')
            break
        else:
            calibration_required = ''
    
    calibration_justification_no = driver.find_element_by_id(
            'txtJustificationNoCal').get_attribute('value')
    calibration_description = driver.find_element_by_id(
            'txtDescriptionofCal').get_attribute('value')
    calibration_frequency = driver.find_element_by_id(
            'txtFrequencyofCal').get_attribute('value')
    
    # Maintenance required - radio button - need to determine if "Yes" or "No"
    maintenance_required_list = driver.find_elements_by_id(
            'rbAssetRequiresScheduledMaintenance')
    for i in maintenance_required_list:
        if i.get_attribute('checked') == 'true':
            maintenance_required = i.get_attribute('value')
            break
        else:
            maintenance_required = ''
    
    maintenance_justification_no = driver.find_element_by_id(
            'txtJustificationNoMaintenance').get_attribute('value')
    maintenance_description = driver.find_element_by_id(
            'txtDescriptionofMaintenance').get_attribute('value')
    maintenance_frequency = driver.find_element_by_id(
            'txtFrequencyofScheduledMaintenance').get_attribute('value')
    
    # Scan attached - radio button - need to determine if "Yes" or "No"
    scan_attached_list = driver.find_elements_by_id('rbInitialCalMaintenance')
    for i in scan_attached_list:
        if i.get_attribute('checked') == 'true':
            scan_attached = i.get_attribute('value')
            break
        else:
            scan_attached = ''
    
    # Get cal/PM file name, if it exists - used later to automatically select if "Scan Attached"
    scan_file = driver.find_element_by_id(
            'mastercontrol.attachments.attachcertificate').text
    
    # Tier level - radio button - need to determine if "Yes" or "No"
    tier_level_list = driver.find_elements_by_id('rbTierLevel')
    for i in tier_level_list:
        if i.get_attribute('checked') == 'true':
            tier_level = i.get_attribute('value')
            break
        else:
            tier_level = ''
    
    # Qualification/validation needed - radio button - need to determine if "Yes" or "No"
    qual_val_needed_list = driver.find_elements_by_id('rbQualVal Needed')
    for i in qual_val_needed_list:
        if i.get_attribute('checked') == 'true':
            qual_val_needed = i.get_attribute('value')
            break
        else:
            qual_val_needed = ''
    
    validation_justification_no = driver.find_element_by_id(
            'txtJustificationNoVal').get_attribute('value')
    validation_description = driver.find_element_by_id(
            'txtBRIEFDescriptionofQual/Val').get_attribute('value')
    
    # Qualification/Validation tab
    qualification_validation_tab = driver.find_element_by_id('formTabsnav3')
    qualification_validation_tab.click()
    
    protocol_number = driver.find_element_by_id(
            'txtProtocolNumber').get_attribute('value')
    comments = driver.find_element_by_id(
            'txtComments').get_attribute('value')
    
    asset_info.extend((asset_number,
                       title,
                       department,
                       manufacturer,
                       manufacturer_model_number,
                       manufacturer_serial_number,
                       vendor,
                       vendor_catalog_number,
                       purchase_order_number,
                       portable,
                       asset_description,
                       intended_use,
                       process_range,
                       classification,
                       new_SOP,
                       new_logbooks,
                       calibration_required,
                       calibration_justification_no,
                       calibration_description,
                       calibration_frequency,
                       maintenance_required,
                       maintenance_justification_no,
                       maintenance_description,
                       maintenance_frequency,
                       scan_attached,
                       scan_file,
                       tier_level,
                       qual_val_needed,
                       validation_justification_no,
                       validation_description,
                       protocol_number,
                       comments))
    
    return asset_info

# Start timer for execution
start_time = time.time()
print('Start time: ' + time.asctime())

# Run function for all asset numbers, log errors
logf = open("error_log.txt", "w")

load_mastercontrol(driver)

for asset_number in range(1900):
    asset_number = str(asset_number).rjust(4, '0')
    try:
        close_extra_windows()
        asset_search(asset_number)
        open_forms_folder()
        expand_search()
        try:
            locate_asset(asset_number)
        except NoAsset:
            continue
        navigate_to_accession_form()
        combined_asset_info.append(gather_data())
        print('Successfully copied asset: ' + asset_number)
    except Exception as e:
        logf.write('{0}:\n{1} {2}\n'.format(str(asset_number),
                   str(traceback.format_exc()), str(e)))

logf.close()

# Write to CSV file
with open('assetInfo_' + str(datetime.datetime.now().strftime('%Y_%m_%d')) +
          '.csv', 'w', newline = '', encoding = 'utf-8') as myfile:
    writer = csv.writer(myfile, delimiter = ',')
    for row in combined_asset_info:
        writer.writerow(row)

# End timer for execution
end_time = time.time()
print('End time: ' + time.asctime())

# Print out total execution time
total_time = end_time - start_time
hours = total_time // 3600
minutes = (total_time - hours * 3600) // 60
seconds = (total_time - hours * 3600 - minutes * 60)

print('Execution time: ' + str(int(hours)) + 'h:' +
      f'{int(minutes):02}' + 'm:' + f'{int(seconds):02}s')