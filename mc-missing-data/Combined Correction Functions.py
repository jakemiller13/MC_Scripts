# -*- coding: utf-8 -*-
"""
Created on Wed Sep 12 12:21:33 2018

@author: jmiller
"""

# Import packages
import pandas as pd
import random
import datetime

# False: print all columns, True: suppress if too large for window
pd.set_option('display.expand_frame_repr', False)

# Generate random number for sampling
rand_num = random.randint(0, 100)

# Load data - MAKE SURE YOU ARE USING MOST RECENT FILE
data = pd.read_csv('assetInfo_2018_12_20.csv', na_values = (''),
                   keep_default_na = False)

# Create N/A list
na_list = ['na', 'Na', 'nA', 'NA', 'n/a', 'N/a', 'n/A', 'N/A']

def missing_info(data):
    '''
    Prints a count of all missing data
    NOTE: This does not count "scan_file" since that is not required field
    '''
    df = data.drop(['scan_file'], axis = 1)
    total = df.size
    blanks = df.isnull().sum().sum()
    na = df.isin(na_list).sum().sum()
    
    # Print data
    print('~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~')
    print(' -------- \n| Totals |\n -------- ')
    print('Accession forms:', str(df.shape[0]).rjust(19))
    print('Total fields counted:', str(total).rjust(14))
    print('Total blanks: ' + (str(blanks) + ', ' +
                              str(round(100 * blanks/total, 2)) +
                              '%').rjust(22))
    print('Total N/A: ' + (str(na) + ', ' +
                           str(round(100 * na/total, 2)) +
                           '%').rjust(25))
    print('~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~')
    print(' ------------ \n| Blank Info |\n ------------ ')
    print(df.isnull().sum())
    print('\n~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~\n')
    print(' ---------------\n| N/A, na, etc. |\n ---------------')
    print(df.isin(na_list).sum())
    print('~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~')

def print_sample(data, rand_num, sample_size):
    '''
    Prints random sample of "sample_size" assets for comparison before/after
    data: loaded above
    rand_num: random number generated to check results before/after
    '''
    print(data.sample(20, random_state = rand_num))

def correct_na(data):
    '''
    Changes any variations of "N/A" to specifically "N/A'
    data: loaded above
    '''
    print(' --------------- ')
    print('| CORRECTING NA |')
    print(' --------------- ')
    
    # Change all versions to 'N/A'
    data.replace(to_replace = na_list, value = 'N/A', inplace = True)

def correct_department(data, department_file):
    '''
    First, updates any missing departments from "department_file"
    Second, changes all departments that are "NaN" to "TBD"
    department_file: csv file which has departments gathered from MC and
                     manually added departments
    data: loaded above
    '''
    print(' -----------------------')
    print('| CORRECTING DEPARTMENT |')
    print(' -----------------------')
    
    departments = pd.read_csv(department_file, encoding = 'cp1252')
    data['department'].fillna(departments['added_department'], inplace = True)
    
def correct_portable(data):
    '''
    Corrects portable assets. These have to be done invidually
    data: loaded above
    '''
    print(' ---------------------')
    print('| CORRECTING PORTABLE |')
    print(' ---------------------')
    
    portable_list = {145 : 'No',
                     146 : 'No',
                     147 : 'Yes',
                     150 : 'No',
                     153 : 'No',
                     155 : 'No',
                     261 : 'No',
                     331 : 'Yes',
                     337 : 'Yes',
                     338 : 'Yes',
                     347 : 'No',
                     383 : 'Yes'}
    
    for asset in portable_list:
        data.loc[(data['asset_number'] == asset), 'portable'] = portable_list[asset]

def correct_calPM_text(data):
    '''
    1.) Checks if calibration/maintenance required is NaN --AND--
        if calibration/maintenance_description or _justification_no is 'N/A'.
        If it is BOTH, assigns "Yes/No" to calibration/maintenance required,
        as appropriate
    2.) Checks if calibration/maintenance is required --AND--
        if corresponding fields are 'N/A'. Inputs 'N/A' only if field is blank
    data: loaded above
    '''    
    print(' ------------------------')
    print('| CORRECTING CAL/PM TEXT |')
    print(' ------------------------')
    
    data.loc[(data['calibration_required'].isnull()) &
             (data['calibration_description'] == 'N/A'),
             'calibration_required'] = 'No'
    data.loc[(data['calibration_required'].isnull()) &
             (data['calibration_justification_no'] == 'N/A'),
             'calibration_required'] = 'Yes'
    data.loc[(data['maintenance_required'].isnull()) &
             (data['maintenance_description'] == 'N/A'),
             'maintenance_required'] = 'No'
    data.loc[(data['maintenance_required'].isnull()) &
             (data['maintenance_justification_no'] == 'N/A'),
             'maintenance_required'] = 'Yes'
             
    data.loc[(data['calibration_required'] == 'Yes') &
             (data['calibration_justification_no'].isnull()),
             'calibration_justification_no'] = 'N/A'
    data.loc[(data['maintenance_required'] == 'Yes') &
             (data['maintenance_justification_no'].isnull()),
             'maintenance_justification_no'] = 'N/A'

    data.loc[(data['calibration_required'] == 'No') &
             (data['calibration_description'].isnull()),
             'calibration_description'] = 'N/A'
    data.loc[(data['maintenance_required'] == 'No') &
             (data['maintenance_description'].isnull()),
             'maintenance_description'] = 'N/A'
    
    # The following are unique and can't be fixed w/ an algorith
#    data.loc[(data['maintenance_required'].isnull() & data['title'].str.\
#              lower().str.contains('timer')), 'maintenance_required'] = 'No'
    data.loc[(data['calibration_required'].isnull()),
             ['calibration_required']] = 'No'
    data.loc[(data['maintenance_required'].isnull()),
             ['maintenance_required']] = 'No'

def correct_calPM_frequency(data):
    '''
    If calibration/maintenance is not required, inserts "N/A" into frequency
    data: loaded above
    tier_level_file: .csv file name in string format
    '''
    print(' -----------------------------')
    print('| CORRECTING CAL/PM FREQUENCY |')
    print(' -----------------------------')
    
    data.loc[(data['calibration_required'] == 'No') |
             (data['calibration_required'] =='FALSE') &
             (data['calibration_frequency'].isnull()),
             'calibration_frequency'] = 'N/A'
    data.loc[(data['maintenance_required'] == 'No') |
             (data['maintenance_required'] =='FALSE') &
             (data['maintenance_frequency'].isnull()),
             'maintenance_frequency'] = 'N/A'

def correct_tier(data, tier_level_file):
    '''
    data: loaded above
    tier_level_file: .csv file name in string format
    '''
    print(' -----------------')
    print('| CORRECTING TIER |')
    print(' ----------------- ')
    
    tiers = pd.read_csv(tier_level_file, encoding = 'cp1252')
    
    count = 0
    
    # Locates asset_number from 'tiers' within 'data' and assigns 'tier_level'
    for i in range(len(tiers)):
        asset_number = tiers.loc[i, 'InfoCard.Dataset..' + 
                                 'prefabReport.InfoCardNumber']
        tier_level = tiers.loc[tiers['InfoCard.Dataset..' +
                                     'prefabReport.InfoCardNumber'] ==
                                     asset_number,
                                     'Asset_Accession_Form_04_23_18__' +
                                     'Tier_Level'].values[0][-1]
        
        try:
            if data.loc[data.asset_number == asset_number,
                        'tier_level'].isnull().item():
                data.loc[data.asset_number == asset_number, 'tier_level'] = \
                        ('Tier ' + tier_level)
                count += 1
        except ValueError:
            pass

def correct_scan_file(data):
    '''
    Checks if there is text in "scan_file" column, assigns yes/no depending
    data: loaded above
    '''    
    print(' ------------------------')
    print('| CORRECTING YES/NO SCAN |')
    print(' ------------------------ ')
    
    data.loc[(data['scan_file'].notnull()), 'scan_attached'] = 'Yes'
    data.loc[(data['scan_file'].isnull()), 'scan_attached'] = \
             'No (Calibration/Maintenance Not Required)'

def correct_validation(data):
    '''
    If qualification/validation is not needed, inputs "N/A" in justification
    data: loaded above
    '''
    print(' ----------------------------')
    print('| CORRECTING VALIDATION INFO |')
    print(' ---------------------------- ')
    
    data.loc[data['qual_val_needed'] == 'No', 'validation_justification_no'] = 'N/A'
    data.loc[data['qual_val_needed'] == 'Yes', 'validation_description'] = 'N/A'
    data.loc[data['qual_val_needed'].isnull(), 'qual_val_needed'] = 'No'

def correct_protocol_number(data):
    '''
    If 'protocol_number' is NaN, inputs 'N/A'
    data: loaded above
    '''
    print(' ----------------------------')
    print('| CORRECTING PROTOCOL NUMBER |')
    print(' ---------------------------- ')
    
    data.loc[(data['protocol_number'].isnull()), 'protocol_number'] = 'N/A'

def correct_new_sop_logbook(data):
    '''
    Changes New SOP/Logbooks to no
    data: loaded above
    '''
    print(' -----------------------------')
    print('| CORRECTING NEW SOP/LOGBOOKS |')
    print(' -----------------------------')
    
    data.loc[data['new_SOP'].isnull(), 'new_SOP'] = 'No'
    data.loc[data['new_logbooks'].isnull(), 'new_logbooks'] = 'No'

def correct_TBD(data):
    '''
    Change any blanks to "TBD" if they are text fields
    These TBD should eventually be figured out
    data: loaded above
    '''
    print(' ------------')
    print('| ADDING TBD |')
    print(' ------------')
    
    list_of_TBD = ['department',
                   'manufacturer',
                   'manufacturer_model_number',
                   'manufacturer_serial_number',
                   'asset_description',
                   'intended_use',
                   'process_range',
                   'calibration_justification_no',
                   'calibration_description',
                   'calibration_frequency',
                   'maintenance_justification_no',
                   'maintenance_description',
                   'maintenance_frequency',
                   'validation_justification_no',
                   'validation_description']
    
    for item in list_of_TBD:
        data.loc[(data[item].isnull()), item] = 'TBD'

def correct_tier_3(data):
    '''
    If Tier Level is NaN for any of the following, changes to "Tier 3":
        -title includes "air handler|
                         condenser|
                         rtu|
                         pump|
                         ph | (Note space after ph to avoid chromatography eg)
                         differential pressure|
                         pressure indicator|
                         manometer|
                         balance|
                         orbital shaker|
                         column"
    '''
    print(' -------------------')
    print('| CORRECTING TIER 3 |')
    print(' -------------------')
    data.loc[(data['title'].str.lower().\
                            str.contains('air handler|'\
                                         'condenser|'\
                                         'rtu|'\
                                         'pump|'\
                                         'ph |'\
                                         'differential pressure|'\
                                         'pressure indicator|'\
                                         'manometer|'\
                                         'balance|'\
                                         'orbital shaker|'\
                                         'column') & \
              data['tier_level'].isnull()), ['tier_level']] = 'Tier 3'
    
    # The following are unique instances and can't be done with algorithm
    data.loc[(data['asset_number'] == 1503), ['tier_level']] = 'Tier 3'

def correct_tier_4(data):
    '''
    If Tier Level is NaN for any of the following, changes to "Tier 4":
        -title includes "timer|
                         pipet|
                         thermometer|
                         chart recorder|
                         incubat| -->covers nonGMP "incubator" and "incubating"
                         wave|
                         peek"
    '''
    print(' -------------------')
    print('| CORRECTING TIER 4 |')
    print(' -------------------')
    data.loc[(data['title'].str.lower().\
                            str.contains('timer|'\
                                         'pipet|'\
                                         'thermometer|'\
                                         'chart recorder|'\
                                         'incubat|'\
                                         'wave|'\
                                         'peek') & \
              data['tier_level'].isnull()), ['tier_level']] = 'Tier 4'
                                         
    # The following are unique instances and can't be done with algorithm
    data.loc[(data['asset_number'].isin([1486,\
                                         1487,\
                                         1488,\
                                         1490,\
                                         1491,\
                                         1492,\
                                         1606])),
                                        ['tier_level']] = 'Tier 4'
        
    # All remaining NaN tier levels were confirmed to not have cal/PM tasks
    data.loc[data['tier_level'].isna(), ['tier_level']] = 'Tier 4'

def correct_classification(data):
    '''
    Corrects any missing classification to GMP
    '''
    data.loc[(data['classification'].isnull()), ['classification']] = 'GMP'

def correct_unnecessary(data):
    '''
    If any of:
            ['title',
            'vendor',
            'vendor_catalog_number',
            'purchase_order_number',
            'calibration_justification_no',
            'maintenance_justification_no']
            are blank (NaN), changes it to 'N/A'
    data: loaded above
    '''
    print(' ------------------------')
    print('| CORRECTING UNNECESSARY |')
    print(' ------------------------')
    
    # Change all NaN values to "N/A"
    data[['title',
          'vendor',
          'vendor_catalog_number',
          'purchase_order_number',
          'calibration_justification_no',
          'maintenance_justification_no',
          'comments']] = \
    data[['title',
          'vendor',
          'vendor_catalog_number',
          'purchase_order_number',
          'calibration_justification_no',
          'maintenance_justification_no',
          'comments']].fillna(value = 'N/A')

def write_missing_tiers_to_csv(data):
    '''
    Writes assets with missing tier levels to CSV, using todays date
    '''
    tier_level_missing = data.loc[(data['tier_level'].isnull()), \
                                 ['asset_number', 'title', 'tier_level']]
    file_name = 'Assets_with_missing_tiers_' + \
                str(datetime.datetime.now().strftime('%Y_%m_%d')) + '.csv'
    tier_level_missing.to_csv(file_name)

def search_titles(data, phrase, missing_tiers = False):
    '''
    Searches data titles for "phrase"
    Returns dataframe of asset_number, title
    '''
    if missing_tiers == True:
        return data.loc[(data['title'].str.lower().str.contains(phrase)),
                        ['asset_number', 'title', 'tier_level']]
    else:
        return data.loc[(data['title'].str.lower().str.contains(phrase)),
                        ['asset_number', 'title']]

# Testing
tier_level_file = 'Tier levels report as of 7.30.18.csv'
department_file = 'Missing Departments.csv'
sample_size = 20

missing_info(data)
#print_sample(data, rand_num, sample_size)
correct_na(data)
correct_department(data, department_file)
correct_portable(data)
correct_calPM_text(data)
correct_calPM_frequency(data)
correct_tier(data, tier_level_file)
correct_new_sop_logbook(data)
correct_scan_file(data)
correct_validation(data)
correct_protocol_number(data)
correct_TBD(data)
correct_tier_3(data)
correct_tier_4(data)
correct_classification(data)
correct_unnecessary(data)
missing_info(data)
#print_sample(data, rand_num, sample_size)

# Save corrected data as csv file
file_name = 'Corrected_Data_' + \
            str(datetime.datetime.now().strftime('%Y_%m_%d')) + '.csv'
data.to_csv(file_name)

##############################################################################
# NOTES:
# 1874 is still being edited. Potentially only run fixing script up to a 
# certain asset number and then stop
##############################################################################