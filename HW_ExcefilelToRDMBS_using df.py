# -*- coding: utf-8 -*-
"""
Created on Sun Apr  8 20:55:00 2018
@author: Amal
"""

# Import `os` 
import os

# Change directory 
os.chdir("C:\\Users\\Amal\\Documents\\python wd")

# Retrieve current working directory (`cwd`)
cwd = os.getcwd()
cwd

import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile

#read all the sheets of the excel files (none=all)
sheets_dict = pd.read_excel('SB_Table1_2001-2009.xlsx', sheet_name=None)

#create an empty data frame
edf = pd.DataFrame()

#loop over the sheets and save the contents and the sheet name in a seperate column
for name, sheet in sheets_dict.items():
    sheet['sheet'] = name
    sheet = sheet.rename(columns=lambda x: x.split('\n')[-1])
    edf = edf.append(sheet)

#reflect the changes on full_table
edf.reset_index(inplace=True, drop=True)

#Set the column labels to equal the values in the 4th row (index location 3):
edf.columns = edf.iloc[3]
edf.rename(columns={'2001':'Year'}, inplace=True)

list(edf.columns.values)
#print(edf )

#drop rows  holding column names
edf = edf[edf.Field != 'Field']

#drop unwanted rows : header part and Totals
edf = edf[~edf['Location'].astype(str).str.contains('Total')]
edf = edf[~edf['Field'].astype(str).str.contains('Total')]
edf = edf[~edf['Field'].astype(str).str.startswith('As of')]
#print(edf)

#drop rows if all rows are nulls ( more than two values in the row)
edf=edf.dropna(thresh=2)        # how='all')        #, inplace=True)
#print(edf)

#fix remaining rows with NaNs, by filling the last filled value of the column
edf=edf.fillna(method='ffill')
#print(edf)

#now the dataframe is ready to copy it to SQL Server
import pyodbc
#create connection string
db_connection = pyodbc.connect(
    Trusted_Connection='Yes',
    Driver='{SQL Server Native Client 11.0}',
    Server='DESKTOP-UR5TAPP',
    Database='ExcelToSqlDB'
)

#######################
import sqlalchemy as sa

engine = sa.create_engine('mssql+pyodbc://DESKTOP-UR5TAPP/ExcelToSqlDB?driver=SQL+Server+Native+Client+11.0')

# write the DataFrame to a table in the sql database
edf.to_sql('SBT1_2001-2009', engine, index=False)

#####################

db_cursor = db_connection.cursor()

### read from sql
stmt = "SELECT * FROM [SBT12001-2009]"
# Excute Query here
dff = pd.read_sql(stmt,db_connection)

dff

