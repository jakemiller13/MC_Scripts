# -*- coding: utf-8 -*-
"""
Created on Tue Aug 21 11:39:04 2018

@author: jmiller
"""

from os import listdir
import shutil
import traceback

folder_path = input('Enter folder location: ')

def split_file_name(file_name):
    '''
    Splits file name from listdir() into parts. File name must be in format: "number_title_date.pdf"
    Returns list of ['file asset number', 'file asset title', 'file calibration date']
    '''
    split_name = file_name[:-4].split('_')
    asset_number = split_name[0]
    file_title = split_name[1]
    calibration_date = split_name[2]
    
    return [asset_number, file_title, calibration_date]

def move_files(folder_path):
    '''
    Moves file from "folder_path" to Cal and PM Certs folder in Engineering
    '''
    for file_name in listdir(folder_path):
        if file_name[-3:] == 'pdf':
            asset_number, file_title, calibration_date = split_file_name(file_name)
            shutil.move(folder_path + '\\' + file_name, 'S:\\Engineering\Cal and PM Certs\\' + asset_number)
            print('Moved: ' + file_name)

logf = open("Moving Files Errors.txt", "w")
try:
    move_files(folder_path)
except Exception as e:
    logf.write('{0}:\n{1}\n'.format(str(traceback.format_exc()), str(e)))