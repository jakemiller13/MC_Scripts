# -*- coding: utf-8 -*-
"""
Created on Wed Aug  8 10:44:40 2018

@author: jmiller
"""

# Import packages
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
import time

# Open DEVELOPMENT MasterControl
url = 'https://acceleronpharma.mastercontrol.com/mc/login/'
driver = webdriver.Chrome(executable_path = r'S:\Engineering\Jake\MasterControl Change CalPM Form\chromedriver')
driver.get(url)

# Setup Waits
wait = WebDriverWait(driver, 5)

# Wait until username/password entered, or 30 seconds
WebDriverWait(driver, 30).until(EC.url_changes(url))

# Navigate to Scheduled Forms
def navigate_to_forms():
    '''Navigates to "Scheduled Forms" from anywhere in MasterControl'''
    wait.until(EC.element_to_be_clickable((By.ID, 'mymastercontrol_text')))
    driver.find_element_by_id('mymastercontrol_text').click()
    wait.until(EC.element_to_be_clickable((By.ID, 'StartTask_name')))
    driver.find_element_by_id('StartTask_name').click()
    wait.until(EC.element_to_be_clickable((By.ID, 'StartFutureForm_text')))
    driver.find_element_by_id('StartFutureForm_text').click()

# Collect number of forms
def total_forms():
    '''Returns total number of forms'''
    try:
        wait.until(EC.frame_to_be_available_and_switch_to_it('myframe'))
    except TimeoutException:
        pass
    number_of_forms = len(driver.find_elements_by_xpath("//tr[starts-with(@id, 'row')]"))
    return number_of_forms

# Open Form by row number
def open_form(row_number):
    '''Opens each form based on row number, NOT on asset number'''
    driver.switch_to_default_content()
    wait.until(EC.frame_to_be_available_and_switch_to_it('myframe'))
    edit_form_button = driver.find_element_by_xpath("//tr[@id = 'row" + str(row_number + 1) + "']/td/a/img[starts-with(@id, 'edit')]")
    edit_form_button.click()

# Replace originator with "name"
def select_name(name):
    '''Selects "name" within "Originator" menu'''
    originator_menu = Select(driver.find_element_by_name('originator'))
    name_selection = driver.find_element_by_xpath("//option[starts-with(@title, '" + name + "')]").text
    originator_menu.select_by_visible_text(name_selection)

# Save Form
def save_form():
    '''Saves form when complete'''
    driver.find_element_by_id('button_save').click()
    driver.switch_to.default_content()
    wait.until(EC.frame_to_be_available_and_switch_to_it('auditFrame'))
    driver.find_element_by_name('Chg_reason').send_keys('Changed originator to individual Process Engineer')
    driver.find_element_by_name('ok').click()

# Edit all Forms
def edit_forms(name):
    '''
    Selects "name" within all forms in "number_of_forms"
    "name" must be string
    Logs errors
    '''
    navigate_to_forms()
    number_of_forms = total_forms()
    
    error_log = open('error_log.txt', 'w')
    
    for i in range(number_of_forms):
        try:
            open_form(i)
            asset_number = driver.find_element_by_name('taskName').get_attribute('value')[0:4]
            select_name(name)
            save_form()
            print('Success editing asset number', asset_number, 'in row', str(i + 1))
        except Exception as e:
            print('Error editing row', str(i))
            error_log.write(e)
    
    error_log.close()

# Start timer for execution
start_time = time.time()
print('Start time: ' + time.asctime())

# Begin execution of script
edit_forms('Jacob Miller')

# End timer for execution
end_time = time.time()
print('End time: ' + time.asctime())

# Print out total execution time
total_time = end_time - start_time
hours = total_time // 3600
minutes = (total_time - hours * 3600) // 60
seconds = (total_time - hours * 3600 - minutes * 60)

print('Execution time: ' + str(int(hours)) + 'h:' + f'{int(minutes):02}' + 'm:' + f'{int(seconds):02}s')