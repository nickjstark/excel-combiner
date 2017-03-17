# ***********************************************************************************************************
# Purpose: Combine multiple Excel sheets (tabs) in to one
# Notes: Checks for duplicates, runs function based on .csv or .xlsx. Also creates index
# Author: Nick Stark
# ***********************************************************************************************************


import subprocess
import pandas as pd
import numpy as np
import os, sys
from os.path import basename

# CSV IMPORT DEFINED FUNCTION
def csvImport(ftype, fpath):
    try:
       if ftype == 1:
           masterdata = pd.read_csv(fpath)
           return masterdata

       if ftype == 2:
           updateddata = pd.read_csv(fpath)
           updateddata['originfile'] = pd.Series(os.path.basename(fpath), \
                                                 index=updateddata.index)
           return updateddata

    except Exception as e:
       print('\nUnable to import CSV file. Error {}'.format(e))
       sys.exit(1)

# EXCEL IMPORT DEFINED FUNCTION
def xlImport(ftype, fpath):
    try:
        if ftype == 1:
           masterdata = pd.read_excel(fpath, 0)
           return masterdata

        if ftype == 2:
           updateddata = pd.read_excel(fpath, 0)
           updateddata['orginfile'] = pd.Series(os.path.basename(fpath), \
                                                index=updateddata.index)
           return updateddata

    except Exception as e:
       print("\nUnable to import Excel file. Error {}".format(e))
       sys.exit(1)

# MASTER FILE USER INPUT DEFINED FUNCTION
def masterfile():
    while True:
       masterfile = input("Enter the path to the master file: ")
       if masterfile.endswith(".csv"):
          return csvImport(1, masterfile)
          break
       elif masterfile.endswith(".xlsx"):
          return xlImport(1, masterfile)
          break
       else:
          print("\nPlease enter a proper CSV format file.")

# UPDATED FILE USER INPUT DEFINED FUNCTION
def updatefile():
    while True:
       updatedfile = input("\nEnter the path to the updated file: ")
       if updatedfile.endswith(".csv"):
          return csvImport(2, updatedfile)
          break
       elif updatedfile.endswith(".xlsx"):
          return xlImport(2, updatedfile)
          break
       else:
          print("\nPlease enter a proper Excel file in xlsx format.")

# CALLING OPENING FUNCTIONS
masterdata = masterfile()
updateddata = updatefile()

# CONCATENATING DATA FRAMES
combineddata = pd.concat([updateddata, masterdata])

# REMOVING DUPLICATES
finaldata = combineddata.drop_duplicates(['Item'])

# SETTING FINAL PATH BY USER INPUT
while True:
    final = input("\nWhere do you want the file, and what do you want to name it? \
                      (e.g., C:\path_to_file\name_of_file.xlsx): ")
    if final.endswith(".xlsx"):
        break
    else:
        print("\nPlease enter a proper Excel file in xlsx format.")

# OUTPUTTING DATA FRAME TO FILE
finaldata.to_excel(final)
print("\nSuccessfully outputted appended data frame to Excel!")

# OPENING OUTPUTTED FILE
# (NOTE: PYTHON STILL RUNS UNTIL SPREADSHEET IS CLOSED)
subprocess.call(final, shell=True)