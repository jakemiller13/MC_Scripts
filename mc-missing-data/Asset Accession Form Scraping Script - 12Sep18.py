#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Jun 13 09:45:52 2018

@author: jmiller
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, WebDriverException, TimeoutException, NoSuchWindowException
import csv
import traceback
import time

combined_asset_info = [['asset_number', 'title', 'department', 'manufacturer', 'manufacturer_model_number', 'manufacturer_serial_number', 'vendor', 'vendor_catalog_number', 'purchase_order_number', 'portable', 'asset_description', 'intended_use', 'process_range', 'classification', 'new_SOP', 'new_logbooks', 'calibration_required', 'calibration_justification_no', 'calibration_description', 'calibration_frequency', 'maintenance_required', 'maintenance_justification_no', 'maintenance_description', 'maintenance_frequency', 'scan_attached', 'scan_file', 'tier_level', 'qual_val_needed', 'validation_justification_no', 'validation_description', 'protocol_number', 'comments']]

# Open MasterControl
url = 'https://acceleronpharma.mastercontrol.com/mc/login/'

##### This is for Windows - uncomment when using work computer #####
driver = webdriver.Chrome(executable_path = r'S:\Engineering\Jake\MasterControl Missing Data\chromedriver')

##### This is for Apple - uncomment when using personal computer #####
#driver = webdriver.Chrome(executable_path = r'/Users/Jake/Documents/Programming/Scraping/chromedriver')

driver.get(url)

# Wait until username/password entered, or 30 seconds
WebDriverWait(driver, 30).until(EC.url_changes(url))

# Search for assets
def assetSearch(asset_number):

    # Switch to main frame
    driver.switch_to.default_content()
    
    search_term = driver.find_element_by_id('strSearchPortal')
    search_term.clear()
    search_term.send_keys(asset_number)
    
    goButton = driver.find_element_by_xpath("//input[@type = 'submit']")
    goButton.click()
    
    # Open "Forms" folder
    driver.switch_to.frame('myframe')
    try:
        driver.execute_script("toggleFolder('image_Forms','group_Forms');")
    except WebDriverException:
        return
    
    # Expand search if >25 records
    try:
        driver.find_element_by_xpath("//div[@id = 'folder_Forms']/a").click()
    except NoSuchElementException:
        pass
    
    # If "Arrow Next" is not "bw", there are more pages we can cycle through
    while True:
    
        # Row may change depending on how many forms associated with asset number
        row_num = 1
        
        while row_num <= 100:
            
            # Try to find asset ID number and click it
            try:
                if str(asset_number) in str(driver.find_element_by_xpath("//div[@id = 'row" + str(row_num) + "']/div[@id = 'column_i.document_num']/a").text):
                    driver.find_element_by_xpath("//div[@id = 'row" + str(row_num) + "']/div[@id = 'column_i.document_num']/a").click()
                    break
            except NoSuchElementException:
                pass
            finally:
                row_num += 1

        # If "Arrow Next" is "bw" and we haven't found asset number yet, break out of While loop
        try:
            if driver.find_element_by_id('viewAdditionalRight').get_attribute('src') == 'https://acceleronpharma.mastercontrol.com/mc/images/icon_arrow_next_bw.gif':
                break
        except NoSuchElementException:
            pass
        
        # If row_num reaches 101, we have reached the end of the page so advance to next page. If we can't advance the page, and we've reached 101, asset doesn't exist so exit function
        try:
            driver.find_element_by_id('viewAdditionalRight').click()
        except NoSuchElementException:
            break

    # Open asset accession form once we've reached it and keep track of windows. Pass if no pop-up. Exit function if we can't find form
    window_before = driver.window_handles[0]
    
    try:
        driver.find_element_by_id('MCD_ALT_TEXT_30').click()
    except NoSuchElementException:
        return
    
    try:
        wait = WebDriverWait(driver, 3)
        wait.until(EC.frame_to_be_available_and_switch_to_it('formFrame'))
    except (TimeoutException, NoSuchWindowException):
        pass
    
    try:
        window_after = driver.window_handles[1]
        driver.switch_to_window(window_after)
    except IndexError:
        pass

    # Gather asset information
    asset_info = []
    
    # General Information tab
    title = driver.find_element_by_id('txtTitle').get_attribute('value')
    
    # Get department info - "selected" doesn't exist if not selected, so need try/except
    try:
        department = driver.find_element_by_xpath("//select[@id = 'mastercontrol.dataset.recordids.Department']/option[@selected = 'true']").text
    except NoSuchElementException:
        department = ''
    
    manufacturer = driver.find_element_by_id('txtManufacturer').get_attribute('value')
    manufacturer_model_number = driver.find_element_by_id('txtManufacturerModelNumber').get_attribute('value')
    manufacturer_serial_number = driver.find_element_by_id('txtManufacturerSerialNumber').get_attribute('value')
    vendor = driver.find_element_by_id('txtVendor').get_attribute('value')
    vendor_catalog_number = driver.find_element_by_id('txtVendorCatalogNumber').get_attribute('value')
    purchase_order_number = driver.find_element_by_id('txtPurchaseOrderNumber').get_attribute('value')

    # Portable - radio button - need to determine if "Yes" or "No"
    portable_list = driver.find_elements_by_id('rbPortable')
    for i in portable_list:
        if i.get_attribute('checked') == 'true':
            portable = i.get_attribute('value')
            break
        else:
            portable = ''
    
    asset_description = driver.find_element_by_id('txtAssetDescription').get_attribute('value')
    intended_use = driver.find_element_by_id('txtIntendedUse').get_attribute('value')
    process_range = driver.find_element_by_id('txtProcessRange').get_attribute('value')

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
    calibration_required_list = driver.find_elements_by_id('rbAssetRequiresScheduledCalibration')
    for i in calibration_required_list:
        if i.get_attribute('checked') == 'true':
            calibration_required = i.get_attribute('value')
            break
        else:
            calibration_required = ''
    
    calibration_justification_no = driver.find_element_by_id('txtJustificationNoCal').get_attribute('value')
    calibration_description = driver.find_element_by_id('txtDescriptionofCal').get_attribute('value')
    calibration_frequency = driver.find_element_by_id('txtFrequencyofCal').get_attribute('value')
    
    # Maintenance required - radio button - need to determine if "Yes" or "No"
    maintenance_required_list = driver.find_elements_by_id('rbAssetRequiresScheduledMaintenance')
    for i in maintenance_required_list:
        if i.get_attribute('checked') == 'true':
            maintenance_required = i.get_attribute('value')
            break
        else:
            maintenance_required = ''
    
    maintenance_justification_no = driver.find_element_by_id('txtJustificationNoMaintenance').get_attribute('value')
    maintenance_description = driver.find_element_by_id('txtDescriptionofMaintenance').get_attribute('value')
    maintenance_frequency = driver.find_element_by_id('txtFrequencyofScheduledMaintenance').get_attribute('value')
    
    # Scan attached - radio button - need to determine if "Yes" or "No"
    scan_attached_list = driver.find_elements_by_id('rbInitialCalMaintenance')
    for i in scan_attached_list:
        if i.get_attribute('checked') == 'true':
            scan_attached = i.get_attribute('value')
            break
        else:
            scan_attached = ''
    
    # Get cal/PM file name, if it exists - used later to automatically select if "Scan Attached"
    scan_file = driver.find_element_by_id('mastercontrol.attachments.attachcertificate').text
    
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
    
    validation_justification_no = driver.find_element_by_id('txtJustificationNoVal').get_attribute('value')
    validation_description = driver.find_element_by_id('txtBRIEFDescriptionofQual/Val').get_attribute('value')
    
    # Qualification/Validation tab
    qualification_validation_tab = driver.find_element_by_id('formTabsnav3')
    qualification_validation_tab.click()
    
    protocol_number = driver.find_element_by_id('txtProtocolNumber').get_attribute('value')
    comments = driver.find_element_by_id('txtComments').get_attribute('value')
    
    asset_info.extend((asset_number, title, department, manufacturer, manufacturer_model_number, manufacturer_serial_number, vendor, vendor_catalog_number, purchase_order_number, portable, asset_description, intended_use, process_range, classification, new_SOP, new_logbooks, calibration_required, calibration_justification_no, calibration_description, calibration_frequency, maintenance_required, maintenance_justification_no, maintenance_description, maintenance_frequency, scan_attached, scan_file, tier_level, qual_val_needed, validation_justification_no, validation_description, protocol_number, comments))
    combined_asset_info.append(asset_info)
    
    # Close new window (if opened) and switch back to main window before exiting function, if necessary
    try:
        driver.switch_to_window(window_after)
        driver.close()
        driver.switch_to_window(window_before)
    except (IndexError, UnboundLocalError):
        pass


# Start timer for execution
start_time = time.time()
print('Start time: ' + time.asctime())

# Run function for all asset numbers, log errors
logf = open("error_log.txt", "w")

for num in range(1900):
    try:
        assetSearch(f'{num:04}')
    except Exception as e:
        logf.write('{0}:\n{1} {2}\n'.format(str(num), str(traceback.format_exc()), str(e)))

logf.close()

# Write to CSV file
with open('combined_asset_info.csv', 'w', newline = '', encoding = 'utf-8') as myfile:
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

print('Execution time: ' + str(int(hours)) + 'h:' + f'{int(minutes):02}' + 'm:' + f'{int(seconds):02}s')