# -*- coding: utf-8 -*-
"""
Created on Thu Aug  9 10:16:53 2018

@author: jmiller
"""

# Import packages
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchWindowException,\
                                       NoSuchElementException,\
                                       UnexpectedAlertPresentException
from selenium.webdriver.common.action_chains import ActionChains
from os import listdir
import time
import traceback
import pandas as pd

# Folder with certs to upload
folder_path = input('Enter folder location: ')

# Read in data file - update file and filepath accordingly
data = pd.read_csv(
       r'S:\\Engineering\\Jake\\MC_Scripts\\mc-missing-data\\' + 
       'Corrected_Data_2018_12_26.csv',
       na_values = (''), keep_default_na = False)

# Define driver and waits
driver = webdriver.Chrome(executable_path =
                          r'S:\\Engineering\\Jake\\MC_Scripts\\' +
                          'mc-missing-data\\chromedriver')
wait = WebDriverWait(driver, 5)

def load_mastercontrol(driver, user_id, password):
    '''
    Opens MasterControl, waits up to 30 seconds for username/password entry
    '''

    url = 'https://acceleronpharma.mastercontrol.com/mc/login/'
    driver.get(url)
    driver.find_element_by_name('userid').send_keys(user_id)
    driver.find_element_by_name('password').send_keys(password)
    driver.find_element_by_id('MCP_BUTTON_40').click()

# Split file name into parts
def split_file_name(file_name):
    '''
    Splits file name from listdir() into parts.
    File name must be in format: "number_title_date.pdf"
    Returns list of ['file asset number',
                     'file asset title',
                     'file calibration date']
    '''
    split_name = file_name[:-4].split('_')
    asset_number = split_name[0]
    file_title = split_name[1]
    calibration_date = format_date(split_name[2])
    
    return [asset_number, file_title, calibration_date]

def format_date(date):
    '''
    Formats "date" into "DD Mmm YYYY" if it is given in some combination of
    "DDMMMYY" or "DD MMM YYYY"
    '''
    day = ''.join([i for i in date[:2] if not i.isalpha()]).replace(' ', '').\
                                                            zfill(2)
    month = ''.join([i for i in date if not i.isdigit()]).replace(' ', '').\
                                                          title()
    year = ''.join([i for i in date[-2:] if not i.isalpha()]).replace(' ', '')
        
    return day + ' ' + month + ' 20' + year

# Get pipette info
def pipette_info(asset_number):
    '''
    Returns: info from data spreadsheet
    Can be used to find specific info using column titles,
        e.g. pipette_info(asset_number).loc['tier_level']
    '''
    info = data.loc[data['asset_number'] == int(asset_number)].iloc[0]
    return info

# Navigate to different frames
def navigate_to_frame(frame = None, window = driver.window_handles[-1]):
    '''
    Navigates between different frames in different windows
    Default is outermost frame in most recently opened window
    frame: None (default), 'myframe', 'formFrame', 'bigFrame'
    window: most recently opened window (default),
        or specify 0-index window handle
    '''
    driver.switch_to_window(window)
    driver.switch_to_default_content()
    wait.until(EC.visibility_of_element_located((By.ID, 'mc')))
    if frame == 'myframe' or frame == 'formFrame':
        wait.until(EC.frame_to_be_available_and_switch_to_it('myframe'))
        if frame == 'formFrame':
            wait.until(EC.frame_to_be_available_and_switch_to_it('formFrame'))
    if frame == 'bigFrame':
        wait.until(EC.frame_to_be_available_and_switch_to_it('bigFrame'))

# Navigate to My Tasks or Scheduled Forms
def navigate_to(location):
    '''
    Navigates to "location" from anywhere in MasterControl
    location: 'My Tasks', 'Scheduled Forms'
    '''
    navigate_to_frame(frame = None, window = driver.window_handles[-1])
    wait.until(EC.element_to_be_clickable((By.ID, 'mymastercontrol_text')))
    driver.find_element_by_id('mymastercontrol_text').click()
    if location.title() == 'My Tasks':
        wait.until(EC.element_to_be_clickable((By.ID, 'MyTasks_name')))
        driver.find_element_by_id('MyTasks_name').click()
        try:
            navigate_to_frame('myframe', window = driver.window_handles[-1])
            Select(driver.find_element_by_name('page')).select_by_value('1')
        except NoSuchElementException:
            pass
    if location.title() == 'Scheduled Forms':
        wait.until(EC.element_to_be_clickable((By.ID, 'StartTask_name')))
        driver.find_element_by_id('StartTask_name').click()
        wait.until(EC.element_to_be_clickable((By.ID, 'StartFutureForm_text')))
        driver.find_element_by_id('StartFutureForm_text').click()

# Check if pipette is already in My Tasks
def check_for_asset(asset_number):
    '''
    Checks if pipette exists.
    If it does, clicks on "Task Name"
    If it does not, will return NoSuchElementException -
        main function navigates to 'Scheduled Forms'
    '''
    navigate_to_frame(frame = 'myframe', window = driver.window_handles[-1])
    driver.find_element_by_xpath(
            "//a[contains(@href, '" + asset_number + "')]").click()

# Complete form if task already exists
def complete_form(asset_number, calibration_date, file_name, user_id,
                  password):
    '''
    Complete form if it exists in 'My Tasks'
    '''
    navigate_to_frame(frame = 'formFrame', window = driver.window_handles[-1])
    driver.execute_script("arguments[0].value = '" + calibration_date + "';",\
                          driver.find_element_by_id('txtDatePerformed_date'))
    ActionChains(driver).move_to_element(driver.find_element_by_xpath(
                         "//input[@id = 'rbInTolerance' and @value = 'Yes']"))\
                         .click().perform()
    ActionChains(driver).move_to_element(driver.find_element_by_xpath(
                         "//input[@id = 'rbExtensionofMetrologyRb' and \
                                  @value = 'No']")).click().perform()
    ActionChains(driver).move_to_element(driver.find_element_by_xpath(
                         "//input[@id = "
                         + "'rbCalandorMaintenanceApprovedandArchivedRb' and "
                         + "@value = 'Yes']")).click().perform()
    attach_pdf(file_name)
    navigate_to_frame(frame = 'myframe', window = driver.window_handles[-1])
    driver.find_element_by_name('signoff').click()
    
    # Sign off form. Check if Alert present for incomplete form.
    #Check if username needed. Log errors
    try:
        navigate_to_frame(frame = 'bigFrame',
                          window = driver.window_handles[-1])
        wait.until(EC.visibility_of_element_located((By.NAME, 'password')))
    except UnexpectedAlertPresentException:   
        driver.switch_to_alert().accept()
    finally:
        try:
            Select(driver.find_element_by_name('status')).\
                   select_by_visible_text('Data Complete')
        except NoSuchElementException:
            pass
        try:
            driver.find_element_by_name('userID').send_keys(user_id)
        except NoSuchElementException:
            pass
        # change password
        driver.find_element_by_name('password').send_keys(password)
        driver.find_element_by_name('frmSave').click()
        
    print(asset_number + ': calibration task successfully COMPLETED from '
          + 'My Tasks with calibration date ' + calibration_date)

# Attach PDF
def attach_pdf(file_name):
    '''
    Attaches pdf to Cal/PM form based on file_name
    '''
    
    driver.find_element_by_name('mastercontrol.attachments.add.'
                                + 'attachedcapprovedvendorcertificate').click()
    
    driver.switch_to_window(driver.window_handles[-1])
    driver.find_element_by_id('MCD_ALT_TEXT_42').click()
    driver.switch_to_window(driver.window_handles[-1])
    driver.find_element_by_name('file_to_upload').\
           send_keys(folder_path + '\\' + file_name)
    driver.execute_script("arguments[0].click()",
                          driver.find_element_by_id('btLoad'))
    driver.switch_to_window(driver.window_handles[-1])
    driver.execute_script("arguments[0].click()",
                          driver.find_element_by_name('buttonPDFForm'))

# Check if pipette asset/PM task exists
def check_for_pipette_cal_task(asset_number):
    '''
    Checks if cal/PM task already exists for "asset_number"
    '''
    navigate_to_frame(frame = 'myframe', window = driver.window_handles[-1])
    driver.find_element_by_xpath("//img[contains(@id, '" + \
                                                 asset_number + "')]")
    print(asset_number + ': has a cal/PM task, but DOES NOT EXIST in My Tasks')

# Room location for pipette
def room_location(department):
    '''
    Department: taken from pipette_info.loc['department']
    Returns room number based on department
    '''
    if department == 'QC':
        return 'Bldg. 128, Room 217/233/237'
    if department == 'BD':
        return 'Bldg. 128, Room 231'
    if department == 'AD':
        return 'Bldg. 128, Room 229'
    if department == 'PD':
        return 'Bldg. 128, Room 213/219/242'
    if department == 'FORM':
        return 'Bldg. 128, Room 215'
    else:
        return 'Portable'

def next_due_date(calibration_date, calibration_frequency,
                  for_initiation = False):
    '''
    Calculates the next calibration due date OR next initiation date for
        scheduling, based on calibration date and frequency
    "calibration_date" is from split_file name
    "calibration_frequency" is "CAL-6, CAL-12, etc."
    for_initiation is False by default. True will return next initiation date
    '''
    months = ['Jan',
              'Feb',
              'Mar',
              'Apr',
              'May',
              'Jun',
              'Jul',
              'Aug',
              'Sep',
              'Oct',
              'Nov',
              'Dec']
    months_till_due = (months.index(''.join(
                       [i for i in calibration_date if not i.isdigit()]).\
                       replace(' ', '')) + int(calibration_frequency.\
                       split('-')[1]))
    if for_initiation == True:
        next_month_due = months[(months_till_due % 12) - 1]
    else:
        next_month_due = months[months_till_due % 12]
    
    next_year_due = str(int(calibration_date[-2:]) + months_till_due//12)
    if for_initiation == True and next_month_due == 'Dec':
        next_year_due = str(int(next_year_due) - 1)
    
    if next_month_due == 'Feb':
        next_day_due = '28'
    elif next_month_due in ['Jan', 'Mar', 'May', 'Jul', 'Aug', 'Oct', 'Dec']:
        next_day_due = '31'
    else:
        next_day_due = '30'
    
    if for_initiation == True:
        next_day_due = '15'
    
    return format_date(next_day_due + next_month_due + next_year_due)


# Create a new Cal/PM task if none exists
def create_new_cal_task(asset_number, calibration_date, due_date,
                        calibration_frequency, maintenance_frequency):
    '''
    Creates new cal/PM task, if none is found
    asset_number, calibration_date, calibration_frequency,
        maintenance_frequency are as expected
    due_date is assumed to be end of month of this calibration, since it's a
        new cal task. Calculated with
        next_due_date(calibration_date, calibration_frequency = 'CAL-0')
    Next due date is calculated from next_due_date()
    '''
    navigate_to('Scheduled Forms')
    navigate_to_frame(frame = 'myframe', window = driver.window_handles[-1])
    driver.find_element_by_name('new').click()
    
    title = asset_number + ': Pipette - Both'
    driver.find_element_by_name('taskName').send_keys(title)  
    
    Select(driver.find_element_by_name('taskRoute')).\
           select_by_visible_text('Asset Calibration / Maint')
    Select(driver.find_element_by_name('originator')).\
           select_by_visible_text('Jacob Miller (JMILLER)')
    driver.find_element_by_name('initiationDate').\
           send_keys(next_due_date(calibration_date, calibration_frequency,
                                   for_initiation = True))
    Select(driver.find_element_by_name('initiationTime')).\
           select_by_visible_text('8 AM')
    driver.find_element_by_name('intervalUnit').\
           send_keys(calibration_frequency.split('-')[1])
    Select(driver.find_element_by_name('interval')).\
           select_by_visible_text('Months')
    
    # Create new Cal/PM Form
    driver.find_element_by_id('portal.scheduling.prepopulate').click()
    
    while True:
        try:
            navigate_to_frame(frame = None, window = driver.window_handles[-1])
            driver.find_element_by_id('mc')
            navigate_to_frame(frame = 'formFrame',
                              window = driver.window_handles[-1])
            break
        
        except (NoSuchWindowException, NoSuchElementException):
            pass
        
    department = pipette_info(asset_number).loc['department']
    
    # Send text info to form
    driver.find_element_by_name('txtTitle').send_keys(title)
    driver.find_element_by_name('txtAssetID').\
           send_keys(str(pipette_info(asset_number).loc['asset_number']).\
           zfill(4))
    driver.find_element_by_name('txtAssetDescription').\
           send_keys(pipette_info(asset_number).loc['asset_description'])
    driver.find_element_by_name('txtModelNumber').\
           send_keys(pipette_info(asset_number).\
           loc['manufacturer_model_number'])
    driver.find_element_by_name('txtSerialNumber').\
           send_keys(pipette_info(asset_number).\
           loc['manufacturer_serial_number'])
    driver.find_element_by_name('txtLocation').\
           send_keys(room_location(department))
    driver.find_element_by_name('txtCalFrequency').\
           send_keys(calibration_frequency)
    driver.find_element_by_name('txtScheduledMaintenanceFrequency').\
           send_keys(maintenance_frequency)   
    
    # Perform dropdown/radio button actions
    Select(driver.find_element_by_name(
           'mastercontrol.dataset.recordids.Department')).\
            select_by_visible_text(department)
    Select(driver.find_element_by_name(
           'mastercontrol.supplier.approved')).\
            select_by_visible_text('Pipette Calibration Services, Inc')
    ActionChains(driver).move_to_element(
                 driver.find_element_by_xpath(
                 "//input[@id = 'rbCalorMaintenance' and @value = 'Both']")).\
                 click().perform()
    ActionChains(driver).move_to_element(
                 driver.find_element_by_xpath(
                 "//input[@id = 'rbIfMaintenancePerformed' and "
                          + "@value = 'Scheduled']")).click().perform()
    
    # Update Calendars - note last performed is same as performed since no cal task had been created previously
    driver.execute_script("arguments[0].value = '" +
                          calibration_date + "';",
                          driver.find_element_by_id(
                                 'txtDateLastPerformed_date'))
    driver.execute_script("arguments[0].value = '" +
                          calibration_date + "';",
                          driver.find_element_by_id(
                                 'txtDatePerformed_date'))
    driver.execute_script("arguments[0].value = '" +
                          format_date(due_date) + "';",
                          driver.find_element_by_id(
                                 'txtDateDue_date'))
    driver.execute_script("arguments[0].value = '" +
                          next_due_date(calibration_date,
                                        calibration_frequency) + "';",
                          driver.find_element_by_id('txtDateNextDue_date'))
    
    # Sign-off and save form
    navigate_to_frame(frame = 'myframe', window = driver.window_handles[-1])
    driver.find_element_by_name('signoff').click()
    wait.until(EC.number_of_windows_to_be, 1)
    navigate_to_frame(frame = 'myframe', window = driver.window_handles[0])
    driver.find_element_by_name('save').click()
    
    print(asset_number + ': new calibration task successfully CREATED '
          + 'with initiation date ' + 
          (next_due_date(calibration_date,
                         calibration_frequency,
                         for_initiation = True)))

def launch_new_cal_task(asset_number):
    '''
    Launches new calibration task
    '''
    navigate_to('Scheduled Forms')
    navigate_to_frame(frame = 'myframe', window = driver.window_handles[-1])
    driver.find_element_by_xpath("//img[contains(@id, '" +
                                                 asset_number + "')]").click()
    
    print(asset_number + ': new calibration task successfully LAUNCHED')

def complete_calibration_tasks(folder_path):
    '''
    Completes all calibration tasks within "folder_path"
    "folder_path" must be in format "DD MMM YY", or some combination like that
    First checks 'My Tasks' to see if calibration task already exists.
        If it exists, completes task.
    If it does not exist, navigates to 'Scheduled Forms' and checks if a
        scheduled form exists.
    If a scheduled form exists, prints to console and moves to next one
        without completing anything
    If a scheduled form does NOT exist, creates new scheduled form
    
    NOTE: "calibration_frequency" and "maintenance_frequency" default to
        CAL-6 and PM-6 if no info available in MasterControl
    '''
    logf = open("error_log.txt", "w")
    for file_name in listdir(folder_path):
        try:
            if file_name[-3:] == 'pdf':
                asset_number, file_title, calibration_date = \
                    split_file_name(file_name)
                
                # Check if "department" exists in for "asset_number" -
                # skip whole section if either "department" or "asset_number"
                # is missing from "data"
                try:
                    check = pipette_info(asset_number).department
                    
                    if pd.isnull(check):
                        print(asset_number +
                              ': DEPARTMENT is missing from info')
                        continue
                    
                except IndexError:
                    print(asset_number +
                          ': does not have any info in ASSET ACCESSION FORM')
                    continue
                
                navigate_to('My Tasks')
                
                while True:
                    try:
                        check_for_asset(asset_number)
                        complete_form(asset_number, calibration_date,
                                      file_name, user_id, password)
                        break
                    
                    except NoSuchElementException:
                        if driver.find_element_by_id('viewAdditionalRight').\
                                  get_attribute('src') == \
                                  'https://acceleronpharma.mastercontrol.com' \
                                  + '/mc/images/icon_arrow_next_bw.gif':
                                
                                navigate_to('Scheduled Forms')
                                
                                try:
                                    check_for_pipette_cal_task(asset_number)
                                
                                except NoSuchElementException:
                                    # These are cal/PM schedules for pipettes
                                    # Change if you do something else
                                    calibration_frequency = 'CAL-6'
                                    maintenance_frequency = 'PM-6'
                                    
                                    due_date = next_due_date(calibration_date,
                                                             'CAL-0')
                                    create_new_cal_task(asset_number,
                                                        calibration_date,
                                                        due_date,
                                                        calibration_frequency,
                                                        maintenance_frequency)
                                    
                                finally:
                                    launch_new_cal_task(asset_number)
                                    navigate_to('My Tasks')
                                
                        else:
                                driver.find_element_by_id(
                                       'viewAdditionalRight').click()
                
                continue
        
        except Exception as e:
            logf.write('{0}:\n{1} {2}\n'.format(str(asset_number),
                       str(traceback.format_exc()), str(e)))

def delete_redundant_tasks(user_id, password):
    '''
    Goes through "My Tasks" and deletes BOLD tasks
    Make sure you've completed these tasks ahead of time.
    This is a CLEAN UP function - if it sees double tasks, it assumes a glitch 
        in the system has created a bunch of extra tasks
    '''
    
    navigate_to('my tasks')
    
    page = 1
    pages = len(Select(driver.find_element_by_name('page')).options)
    
    while page <= pages:
        navigate_to('my tasks')
        
        try:
            Select(driver.find_element_by_name('page')).\
                   select_by_value(str(page))
        except NoSuchElementException:
            print('Page does not exist')
        
        row = 1
        
        while row <= 100:
            print('row: ' + str(row))
            try:
                print('here 1')
                if 'bold' in driver.find_element_by_id('row' + str(row)).\
                                                       get_attribute('style'):
                    print('here 2')
                    driver.find_element_by_xpath("//\
                                                 div[@id = 'row" + str(row) \
                                                     + "']/\
                                                 div[@id = 'column_actions']/\
                                                 a[@name = '_Abort']").click()
            
                    driver.switch_to.alert.accept()
                    navigate_to_frame(frame = 'bigFrame',
                                      window = driver.window_handles[-1])
                    driver.find_element_by_name('comments').\
                           send_keys('redundant task created by system')
                    
                    try:
                        driver.find_element_by_name('userID').\
                                                    send_keys(user_id)
                    except NoSuchElementException:
                        pass
                    
                    driver.find_element_by_name('password').send_keys(password)
                    driver.find_element_by_name('frmSave').click()
                    navigate_to_frame('myframe')
                    
                else:
                    print('here 3')
                    row += 1
            
            except NoSuchElementException:
                print('here 4')
                row += 1
                pass

        page += 1

def delete_pipettes(user_id, password):
    '''
    Goes through "My Tasks" and checks for anything with "Pipette" in title
    Make sure you've completed these tasks ahead of time!
    This is a CLEAN UP function - if it sees "Pipette", it gets deleted
    '''
    
    navigate_to('my tasks')
    
    page = 1
    pages = len(Select(driver.find_element_by_name('page')).options)
    
    while page <= pages:
        navigate_to('my tasks')
        Select(driver.find_element_by_name('page')).select_by_value(str(page))

        count = len(driver.find_elements_by_xpath(
                    "//img[contains(@id, 'abort_') and "\
                           + "contains(@id, 'pette')]"))
            
        while count > 0:
            print('Remaining tasks: ' + str(count))
            
            navigate_to_frame('myframe')
                
            driver.find_element_by_xpath("//img[contains(@id, 'abort_') and "
                                                + "contains(@id, 'pette')]").\
                                         click()
            driver.switch_to.alert.accept()
            navigate_to_frame(frame = 'bigFrame',
                              window = driver.window_handles[-1])
            driver.find_element_by_name('comments').\
                   send_keys('redundant task created by system')
                
            try:
                driver.find_element_by_name('userID').send_keys(user_id)
            except NoSuchElementException:
                pass
                
            driver.find_element_by_name('password').send_keys(password)
            driver.find_element_by_name('frmSave').click()

        count = len(driver.find_elements_by_xpath(
                    "//img[contains(@id, 'abort_') and "\
                           + "contains(@id, 'pette')]"))

        navigate_to_frame('myframe')
        
        if driver.find_element_by_id('viewAdditionalRight').\
                  get_attribute('src') == \
                  'https://acceleronpharma.mastercontrol.com/mc/images' \
                  + '/icon_arrow_next.gif':
            page += 1
            driver.find_element_by_id('viewAdditionalRight').click()
        else:
            break

start_time = time.time()
print('Start time: ' + time.asctime())

user_id = input('User Name: ')
password = input('Password: ')

load_mastercontrol(driver, user_id, password)
print('You are now free to use whatever functions you need')
end_time = time.time()
print('End time: ' + time.asctime())

total_time = end_time - start_time
hours = total_time // 3600
minutes = (total_time - hours * 3600) // 60
seconds = (total_time - hours * 3600 - minutes * 60)

print('Execution time: ' +
      str(int(hours)) + 'h:' +
      f'{int(minutes):02}' + 'm:' +
      f'{int(seconds):02}s')