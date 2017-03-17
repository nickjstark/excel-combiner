# ***************************************************
# Purpose: Read in Excel sheets as dataframes
# Author: Nick Stark
# ***************************************************

import pandas as pd

df1 = pd.read_excel("T:\Maintenance\Planner Stuff\Activities\API ACTIVITY DATA.xlsx",
                    sheetname="ACTIVITY HEADER", header=0, parse_cols="A, C")

# DEBUG PRINT
# print(df1)

df2 = pd.read_excel("T:\Maintenance\Planner Stuff\Activities\API METER & PM DATA.xlsx",
                    sheetname="DATE PM'S", header=0, parse_cols="A, F, K")
df2 = df2.fillna(method='ffill')

# DEBUG PRINT
# print(df2)

# PERFORM JOIN TO COMBINE THE DATAFRAMES
merge = pd.merge(df2, df1, how="left", on="ACTIVITY NAME", sort=True, indicator=True)

# DEBUG PRINT
# print(merge)
