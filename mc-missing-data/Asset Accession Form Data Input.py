# -*- coding: utf-8 -*-
"""
Created on Tue Jul 24 08:57:06 2018

@author: jmiller
"""

# Import packages
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, WebDriverException, TimeoutException, NoSuchWindowException
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
import traceback
import time


# Read in data file
data = pd.read_csv('Original Asset Info.csv', na_values = (''), keep_default_na = False)

# Open MasterControl
url = 'https://acceleronpharma.mastercontrol.com/mcdev/login/'

##### This is for Windows - uncomment when using work computer #####
driver = webdriver.Chrome(executable_path = r'S:\Engineering\Jake\MasterControl Missing Data\chromedriver')

##### This is for Apple - uncomment when using personal computer #####
#driver = webdriver.Chrome(executable_path = r'/Users/Jake/Documents/Programming/Scraping/chromedriver')

driver.get(url)

###### Uncomment this when you are actually executing #####
#
## Wait until username/password entered, or 30 seconds
#WebDriverWait(driver, 30).until(EC.url_changes(url))
#
###### Uncomment this when you are actually executing #####

# Locate Asset Accession form
def locate(asset_number):
    
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
        
    # Check out form for revision, or log error if form already checked out
    driver.find_element_by_id('button_revision').click()
    try:
        driver.find_element_by_link_text('Check Out').click()
        driver.find_element_by_link_text('Check Out').click()
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
    
    # Send info from data file to correct field
    
    ##### TODO: if statements for if field already contains info
    ##### TODO: try to close out second window and switch back to first window
    ##### TODO: complicated, but can you figure out a for loop that excludes radio button values?
    
    # General Information tab
    driver.find_element_by_id('txtTitle').clear()
    driver.find_element_by_id('txtTitle').send_keys(data[data.asset_number == int(asset_number)].title.values[0])
    
    # Department - dropdown menu
    department = Select(driver.find_element_by_xpath("//select [@id = 'mastercontrol.dataset.recordids.Department']"))
    department.select_by_value(data[data.asset_number == int(asset_number)].department.values[0])
    
    driver.find_element_by_id('txtManufacturer').clear()
    driver.find_element_by_id('txtManufacturer').send_keys(data[data.asset_number == int(asset_number)].manufacturer.values[0])
    
    driver.find_element_by_id('txtManufacturerModelNumber').clear()
    driver.find_element_by_id('txtManufacturerModelNumber').send_keys(data[data.asset_number == int(asset_number)].manufacturer_model_number.values[0])
    
    driver.find_element_by_id('txtManufacturerSerialNumber').clear()
    driver.find_element_by_id('txtManufacturerSerialNumber').send_keys(data[data.asset_number == int(asset_number)].manufacturer_serial_number.values[0])
    
    driver.find_element_by_id('txtVendor').clear()
    driver.find_element_by_id('txtVendor').send_keys(data[data.asset_number == int(asset_number)].vendor.values[0])
    
    driver.find_element_by_id('txtVendorCatalogNumber').clear()
    driver.find_element_by_id('txtVendorCatalogNumber').send_keys(data[data.asset_number == int(asset_number)].vendor_catalog_number.values[0])
    
    driver.find_element_by_id('txtPurchaseOrderNumber').clear()
    driver.find_element_by_id('txtPurchaseOrderNumber').send_keys(data[data.asset_number == int(asset_number)].purchase_order_number.values[0])

    # Portable - radio button
    portable_yes_no = data[data.asset_number == int(asset_number)].portable.values[0]
    portable_button = driver.find_element_by_xpath("//input[@id = 'rbPortable' and @value = '" + portable_yes_no + "']")
    ActionChains(driver).move_to_element(portable_button).click().perform()
    
    driver.find_element_by_id('txtAssetDescription').clear()
    driver.find_element_by_id('txtAssetDescription').send_keys(data[data.asset_number == int(asset_number)].asset_description.values[0])
    
    driver.find_element_by_id('txtIntendedUse').clear()
    driver.find_element_by_id('txtIntendedUse').send_keys(data[data.asset_number == int(asset_number)].intended_use.values[0])
    
    driver.find_element_by_id('txtProcessRange').clear()
    driver.find_element_by_id('txtProcessRange').send_keys(data[data.asset_number == int(asset_number)].process_range.values[0])
    
    # Classification - radio button
    classification_value = data[data.asset_number == int(asset_number)].classification.values[0]
    classification_button = driver.find_element_by_xpath("//input[@id = 'rbClassification' and @value = '" + classification_value + "']")
    ActionChains(driver).move_to_element(classification_button).click().perform()
    
    # New SOP - radio button
    new_SOP_yes_no = data[data.asset_number == int(asset_number)].new_SOP.values[0]
    new_SOP_button = driver.find_element_by_xpath("//input[@id = 'rbYes/No3' and @value = '" + new_SOP_yes_no + "']")
    ActionChains(driver).move_to_element(new_SOP_button).click().perform()
    
    # New logbooks - radio button
    new_logbooks_yes_no = data[data.asset_number == int(asset_number)].new_logbooks.values[0]
    new_logbooks_button = driver.find_element_by_xpath("//input[@id = 'rbYes/No4' and @value = '" + new_logbooks_yes_no + "']")
    ActionChains(driver).move_to_element(new_logbooks_button).click().perform()
    
    # Switch to Calibration & Maintenance tab
    driver.find_element_by_id('formTabsnav2').click()
    
    # Calibration required - radio button
    calibration_yes_no = data[data.asset_number == int(asset_number)].calibration_required.values[0]
    calibration_button = driver.find_element_by_xpath("//input[@id = 'rbAssetRequiresScheduledCalibration' and @value = '" + calibration_yes_no + "']")
    ActionChains(driver).move_to_element(calibration_button).click().perform()
    
    driver.find_element_by_id('txtJustificationNoCal').clear()
    driver.find_element_by_id('txtJustificationNoCal').send_keys(data[data.asset_number == int(asset_number)].calibration_justification_no.values[0])
    
    driver.find_element_by_id('txtDescriptionofCal').clear()
    driver.find_element_by_id('txtDescriptionofCal').send_keys(data[data.asset_number == int(asset_number)].calibration_description.values[0])
    
    driver.find_element_by_id('txtFrequencyofCal').clear()
    driver.find_element_by_id('txtFrequencyofCal').send_keys(data[data.asset_number == int(asset_number)].calibration_frequency.values[0])
    
    # Maintenance required - radio button
    maintenance_yes_no = data[data.asset_number == int(asset_number)].maintenance_required.values[0]
    maintenance_button = driver.find_element_by_xpath("//input[@id = 'rbAssetRequiresScheduledMaintenance' and @value = '" + maintenance_yes_no + "']")
    ActionChains(driver).move_to_element(maintenance_button).click().perform()
    
    driver.find_element_by_id('txtJustificationNoMaintenance').clear()
    driver.find_element_by_id('txtJustificationNoMaintenance').send_keys(data[data.asset_number == int(asset_number)].maintenance_justification_no.values[0])
    
    driver.find_element_by_id('txtDescriptionofMaintenance').clear()
    driver.find_element_by_id('txtDescriptionofMaintenance').send_keys(data[data.asset_number == int(asset_number)].maintenance_description.values[0])
    
    driver.find_element_by_id('txtFrequencyofScheduledMaintenance').clear()
    driver.find_element_by_id('txtFrequencyofScheduledMaintenance').send_keys(data[data.asset_number == int(asset_number)].maintenance_frequency.values[0])
    
    # Scan attached - radio button
    scan_attached_yes_no = data[data.asset_number == int(asset_number)].scan_attached.values[0]
    scan_attached_button = driver.find_element_by_xpath("//input[@id = 'rbInitialCalMaintenance' and @value = '" + scan_attached_yes_no + "']")
    ActionChains(driver).move_to_element(scan_attached_button).click().perform()
    
    # Tier level - radio button
    tier_level_value = data[data.asset_number == int(asset_number)].tier_level.values[0]
    tier_level_button = driver.find_element_by_xpath("//input[@id = 'rbTierLevel' and @value = '" + tier_level_value + "']")
    ActionChains(driver).move_to_element(tier_level_button).click().perform()
    
    # Qualification/validation needed - radio button
    qual_val_needed_yes_no = data[data.asset_number == int(asset_number)].qual_val_needed.values[0]
    qual_val_needed_button = driver.find_element_by_xpath("//input[@id = 'rbQualVal Needed' and @value = '" + qual_val_needed_yes_no + "']")
    ActionChains(driver).move_to_element(qual_val_needed_button).click().perform()
    
    driver.find_element_by_id('txtJustificationNoVal').clear()
    driver.find_element_by_id('txtJustificationNoVal').send_keys(data[data.asset_number == int(asset_number)].validation_justification_no.values[0])
    
    driver.find_element_by_id('txtBRIEFDescriptionofQual/Val').clear()
    driver.find_element_by_id('txtBRIEFDescriptionofQual/Val').send_keys(data[data.asset_number == int(asset_number)].validation_description.values[0])
    
    # Switch to Qualification/Validation tab
    driver.find_element_by_id('formTabsnav3').click()
    
    driver.find_element_by_id('txtProtocolNumber').clear()
    driver.find_element_by_id('txtProtocolNumber').send_keys(data[data.asset_number == int(asset_number)].protocol_number.values[0])
    
    driver.find_element_by_id('txtComments').clear()
    driver.find_element_by_id('txtComments').send_keys(data[data.asset_number == int(asset_number)].comments.values[0])


##### EXECUTION + TESTING #####

# Start timer for execution
start_time = time.time()
print('Start time: ' + time.asctime())

##### ----> TESTING BEGINS HERE <----- #####

logf = open("error_log.txt", "w")

for num in range(1820,1821):
    try:
        locate(f'{num:04}')
    except Exception as e:
        logf.write('{0}:\n{1} {2}\n'.format(str(num), str(traceback.format_exc()), str(e)))

logf.close()

##### ----> TESTING ENDS HERE <----- #####

# End timer for execution
end_time = time.time()
print('End time: ' + time.asctime())

# Print out total execution time
total_time = end_time - start_time
hours = total_time // 3600
minutes = (total_time - hours * 3600) // 60
seconds = (total_time - hours * 3600 - minutes * 60)

print('Execution time: ' + str(int(hours)) + 'h:' + f'{int(minutes):02}' + 'm:' + f'{int(seconds):02}s')