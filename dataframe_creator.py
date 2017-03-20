# ***************************************************
# Purpose: Read in Excel sheets as dataframes
#          Output dataframes to a new Excel file
# Author: Nick Stark
# ***************************************************

import pandas as pd

df1 = pd.read_excel("T:\Maintenance\Planner Stuff\Activities\API ACTIVITY DATA.xlsx",
                    sheetname="ACTIVITY HEADER", header=0, parse_cols="A")

# DEBUG PRINT
# print(df1)

df2 = pd.read_excel("T:\Maintenance\Planner Stuff\Activities\API METER & PM DATA.xlsx",
                    sheetname="DATE PM'S", header=0, parse_cols="A, F, K")
df2 = df2.fillna(method='ffill')

# DEBUG PRINT
# print(df2)

df3 = pd.read_excel("T:\Maintenance\Planner Stuff\Activities\API ACTIVITY DATA.xlsx",
                    sheetname="ACTIVITY ROUTING", header=0, parse_cols="A, F")
df3 = df3.fillna(method='ffill')

# PERFORM JOIN TO COMBINE THE DATAFRAMES
merge = pd.merge(df2, df1, how="left", on="ACTIVITY NAME", sort=True, indicator=False)
merge = merge.sort_values(["ASSET NUMBER", "ACTIVITY NAME"])

# DEBUG PRINT
# print(merge)

merge2 = pd.merge(df3, merge, how='left', on="ACTIVITY NAME", sort=True, indicator=True)
merge2 = merge2.sort_values(['ASSET NUMBER', 'ACTIVITY NAME', 'RESOURCE'])

# DEBUG PRINT
# print(merge2)

# INITIALIZE EXCELWRITER AND OUTPUT FILE
writer = pd.ExcelWriter('T:\Maintenance\Planner Stuff\Activities\PM_INFO.xlsx', engine='xlsxwriter')
merge2.to_excel(writer, 'PM_INFO')
writer.save()
