# -*- coding: utf-8 -*-
"""
Created on Tue Jul 31 13:15:53 2018

@author: MANHARDTD
"""

from openpyxl import load_workbook
from pandas import DataFrame
import os

#create tables to store sig wave height values to add to dataframe and output to excel sheet
file_store = []
toe_station_store = []
toe_elevation_store = []
top_station_store = []
top_elevation_store = []

#takes parameters based on method used in runup and outputs a "Values.xlsx" file containing
#the significant wave height and wave period for each runup .xlsm file in the directory specified

def FindStructureInfo(toe_station_cell, toe_elevation_cell, top_station_cell, top_elevation_cell):

    sheet_ranges = wb['Profile Points']
    toe_station = (sheet_ranges[toe_station_cell].value)
    toe_elevation = (sheet_ranges[toe_elevation_cell].value)
    top_station = (sheet_ranges[top_station_cell].value)
    top_elevation = (sheet_ranges[top_elevation_cell].value)
        
    #write values to excel file
    file_store.append(fileToOpen)
    toe_station_store.append(toe_station)
    toe_elevation_store.append(toe_elevation)
    top_station_store.append(top_station)
    top_elevation_store.append(top_elevation)

#Path where excel files are located, needs to be updated for each county
    
directory = input("What is the directory where the runup files are stored? ")
directory = os.chdir(directory)
#loop through each file in the directory and opens file
for file in os.listdir(directory):
    filename = os.fsdecode(file)
    if filename.endswith('.xlsm') and "~$" not in filename:
        fileToOpen = filename
        os.getcwd()
        print(fileToOpen)
        wb = load_workbook(filename = fileToOpen, data_only = True)
        print("Loaded file")

#check what method (type 1-3) and update parameters for that method, then run FindSigWaveValues()

        if "Type 1_Stockdon_Erosion" in fileToOpen:
            toe_station_cell = "ZS2"
            toe_elevation_cell = "ZS3"
            top_station_cell = "ZS4"
            top_elevation_cell = "ZS5"
        
            FindStructureInfo(toe_station_cell, toe_elevation_cell, top_station_cell, top_elevation_cell)
        
        elif "Type 1" in fileToOpen:
            toe_station_cell = "CS2"
            toe_elevation_cell = "CS3"
            top_station_cell = "CS4"
            top_elevation_cell = "CS5"
        
            FindStructureInfo(toe_station_cell, toe_elevation_cell, top_station_cell, top_elevation_cell)
            
        #if type 2 method cells change
        elif "Type 2" in fileToOpen:
            print("opened file as type 2")
            toe_station_cell = "G2"
            toe_elevation_cell = "G3"
            top_station_cell = "G4"
            top_elevation_cell = "G5"
        
            FindStructureInfo(toe_station_cell, toe_elevation_cell, top_station_cell, top_elevation_cell)
        
        #if type 3 method, cells and sheets change
        elif "Type 3" in fileToOpen:
            print("opened file as type 3")
            toe_station_cell = "F2"
            toe_elevation_cell = "F3"
            top_station_cell = "F4"
            top_elevation_cell = "F5"
        
            FindStructureInfo(toe_station_cell, toe_elevation_cell, top_station_cell, top_elevation_cell)
        else:
            print("ERROR")
            quit()
    #write values to excel file
    df = DataFrame({'File': file_store, 'Toe Station': toe_station_store, 'Toe Elevation': toe_elevation_store, 'Top Station': top_station_store, 'Top Elevation': top_elevation_store})
    print(df)
    df.to_excel('StructureInfo.xlsx', sheet_name = 'sheet1', index = False)