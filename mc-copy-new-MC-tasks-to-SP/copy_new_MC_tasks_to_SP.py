# -*- coding: utf-8 -*-
"""
Created on Fri Sep 21 15:47:38 2018

@author: jmiller
"""

from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import time

# Websites --- REMEMBER IF YOU ARE WORKING IN DEV OR NOT
mc_url = 'https://acceleronpharma.mastercontrol.com/mc/login/'
#mc_url = 'https://acceleronpharma.mastercontrol.com/mcdev/login/'
sp_url = 'https://xlrn.sharepoint.com/TechOps/Manufacturing/Drug_Product/' + \
         'Engineering/Lists/Upcoming%20CalibrationMaintenance/' + \
         'Not%20completed.aspx'

class Driver:
    def __init__(self):
        self.driver = webdriver.Chrome(executable_path = \
                                       r'S:\Engineering\Jake\chromedriver')
        self.wait = WebDriverWait(self.driver, 5)


class MC_Driver(Driver):
    '''
    Class for manipulating MasterControl
    '''

    def load_mastercontrol(self, url):
        '''
        Opens MasterControl, waits up to 30 seconds for username/password entry
        '''
        self.driver.get(url)
        WebDriverWait(self.driver, 30).until(EC.url_changes(url))
    
    def navigate_to_frame(self, frame = None):
        '''
        Navigates between different frames
        Default is outermost frame in most recently opened window
        frame: None (default), 'myframe', 'formFrame', 'bigFrame'
        '''
        self.driver.switch_to_window(self.driver.window_handles[-1])
        self.driver.switch_to_default_content()
        self.wait.until(EC.visibility_of_element_located((By.ID, 'mc')))
        if frame == 'myframe' or frame == 'formFrame':
            self.wait.until(EC.frame_to_be_available_and_switch_to_it(
                            'myframe'))
            if frame == 'formFrame':
                self.wait.until(EC.frame_to_be_available_and_switch_to_it(
                                'formFrame'))
        if frame == 'bigFrame':
            self.wait.until(EC.frame_to_be_available_and_switch_to_it(
                            'bigFrame'))

    def navigate_to(self, location):
        '''
        Navigates to "location" from anywhere in MasterControl
        location: 'My Tasks', 'Scheduled Forms'
        '''
        self.navigate_to_frame(frame = None)
        self.wait.until(EC.element_to_be_clickable((By.ID,
                                                    'mymastercontrol_text')))
        self.driver.find_element_by_id('mymastercontrol_text').click()
        if location.title() == 'My Tasks':
            self.wait.until(EC.element_to_be_clickable((By.ID,
                                                        'MyTasks_name')))
            self.driver.find_element_by_id('MyTasks_name').click()
            try:
                self.navigate_to_frame('myframe')
                Select(self.driver.find_element_by_name('page')).\
                       select_by_value('1')
            except NoSuchElementException:
                pass
        if location.title() == 'Scheduled Forms':
            self.wait.until(EC.element_to_be_clickable(
                           (By.ID, 'StartTask_name')))
            self.driver.find_element_by_id('StartTask_name').click()
            self.wait.until(EC.element_to_be_clickable(
                           (By.ID, 'StartFutureForm_text')))
            self.driver.find_element_by_id('StartFutureForm_text').click()
    
    def check_for_new(self, row_num):
        '''
        Checks row1 to row100 if style contains "bold"
        If it does, opens task and runs get_info()
        Then navigates back to "My Tasks"
        '''
        if 'bold' in self.driver.find_element_by_id('row' + str(row_num)).\
                                                    get_attribute('style'):
            return True
    
    def open_task(self, row_num):
        '''
        Opens a task if check_for_new() finds a bold task
        Gets info using get_info
        Returns to "My Tasks"
        '''
        self.driver.find_element_by_xpath("//div[@id = 'row" + \
                                          str(row_num) + "']/div[@id = " + \
                                          "'column_packet_nm']/a").click()
    
    def get_info(self):
        '''
        Gets relevant info from MasterControl and converts to SharePoint format
        '''
        self.navigate_to_frame('formFrame')
        
        title = self.driver.find_element_by_id('txtTitle').\
                                               get_attribute('value')
        department = Select(self.driver.find_element_by_id(\
                            'mastercontrol.dataset.recordids.Department')).\
                            first_selected_option.text
        due_date = self.driver.find_element_by_id('txtDateDue_date').\
                                                  get_attribute('value')
        performed_by = Select(self.driver.find_element_by_id(\
                              'mastercontrol.supplier.approved')).\
                              first_selected_option.text
        
        sp_department = dept_conversion(department)
        sp_due_date = date_conversion(due_date)
        
        return title, sp_department, sp_due_date, performed_by
    
    def select_page(self, page_num):
        '''
        Switches to selected page, if available
        Returns "break" if no more pages to select
        '''
        Select(self.driver.find_element_by_name('page')).\
               select_by_value(str(page_num))


class SP_Driver(Driver):
    '''
    Class for manipulating SharePoint
    '''
    
    def load_sharepoint(self, url):
        '''
        Opens SharePoint, waits up to 30 seconds for username/password entry
        Note that this function uses a different url and waits for it to MATCH.
        May have to update if that url ever changes
        '''
        self.driver.get(url)
        WebDriverWait(self.driver, 30).until(EC.url_matches(\
                     'https://xlrn.sharepoint.com/TechOps/Manufacturing/' + \
                     'Drug_Product/Engineering/Lists/' + \
                     'Upcoming%20CalibrationMaintenance/Not%20completed.aspx'))
    
    def add_task(self):
        '''
        Opens new SharePoint task
        '''
        self.driver.find_element_by_id('idHomePageNewItem').click()
    
    def send_info(self, title, sp_department, sp_due_date, performed_by):
        '''
        Sends:
            "title" to "Task Name"
            "sp_department" to "Equipment Owner"
            "sp_due_date" to "Due Date"
            "performed_by" to "Additional info"
        '''
        self.driver.find_element_by_xpath('//input[@title =' + \
                                          '"Task Name Required Field"]').\
                                          send_keys(title)
        self.driver.find_element_by_xpath('//input[@title =' + \
                                          '"Equipment Owner"]').\
                                          send_keys(sp_department)
        self.driver.find_element_by_xpath('//input[@title =' + \
                                          '"Due Date"]').\
                                          send_keys(sp_due_date)
        self.driver.find_element_by_id(
            'Body_7662cd2c-f069-4dba-9e35-082cf976e170_$TextField_inplacerte')\
                .send_keys(performed_by)
    
    def save_task(self):
        '''
        Saves SharePoint task
        Separate function for troubleshooting (javascript)
        '''
        self.driver.find_element_by_xpath('//li[contains(@id,' + \
                                          '"Ribbon.ListForm.Edit-title")]').\
                                          click()
        self.driver.find_element_by_xpath('//a[contains(@id,' + \
                                          '"Ribbon.ListForm.Edit.' + \
                                          'Commit.Publish")]').\
                                          click()


# Generic (i.e. classless) functions
def startup():
    mc = MC_Driver()
    sp = SP_Driver()
    mc.load_mastercontrol(mc_url)
    sp.load_sharepoint(sp_url)
    return mc, sp

def date_conversion(date):
    '''
    Separate function to convert dates between MasterControl and SharePoint
    Currently set up to convert from MasterControl to SharePoint
    '''
    months = {'Jan':'1',
              'Feb':'2',
              'Mar':'3',
              'Apr':'4',
              'May':'5',
              'Jun':'6',
              'Jul':'7',
              'Aug':'8',
              'Sep':'9',
              'Oct':'10',
              'Nov':'11',
              'Dec':'12'}
    
    if ' ' in date:
        day = date[0:2]
        month = months[date[3:6]]
        year = date[-4:]
        sp_due_date = month + '/' + day + '/' + year
        return sp_due_date

def dept_conversion(department):
    '''
    Separate function to convert departments between MasterControl and
        SharePoint
    Currently set up to convert from MasterControl to SharePoint
    '''
    dept_mc_to_sp = {'MF':'MSAT TEAM',
                     'QC':'QC',
                     'AD':'Analytical Members',
                     'EN':'Engineering Members',
                     'MM':'Materials Management'}
    
    try:
        return dept_mc_to_sp[department]
    except KeyError:
        return ''

def run_program(program_name, error_log_name):
    '''
    Runs "program_name"
    Records errors
    Prints out total time upon completion
    '''
    start_time = time.time()
    print('Start time: ' + time.asctime())
    
    logf = open(error_log_name, "w")
    try:
        program_name()
    except Exception as e:
        logf.write(str(e))
    
    end_time = time.time()
    print('End time: ' + time.asctime())
    
    total_time = end_time - start_time
    hours = total_time // 3600
    minutes = (total_time - hours * 3600) // 60
    seconds = (total_time - hours * 3600 - minutes * 60)
    
    print('Execution time: ' + str(int(hours)) + 'h:' + \
          f'{int(minutes):02}' + 'm:' + \
          f'{int(seconds):02}s')

# Main program
def MC_to_SP():
    mc, sp = startup()
    mc.navigate_to('my tasks')
    mc.navigate_to_frame('myframe')
    page_num = 1
    try:
        while True:
            row_num = 1
            mc.select_page(page_num)
            while row_num <= 100:
                if mc.check_for_new(row_num):
                    mc.open_task(row_num)
                    sp.add_task()
                    sp.send_info(*mc.get_info())
                    sp.save_task()
                    print('"' + mc.get_info()[0] + '" successfully copied')
                    mc.navigate_to('my tasks')
                    mc.select_page(page_num)
                    mc.navigate_to_frame('myframe')
                row_num += 1
            page_num += 1
    except NoSuchElementException:
        pass

run_program(MC_to_SP, 'MC_to_SP_errorlog.txt')