import pandas as pd
import xlrd
import sqlalchemy
import math
from sqlalchemy import create_engine
from pandas_profiling import ProfileReport
from snowflake.connector.pandas_tools import write_pandas
import snowflake.connector
import numpy as np
import win32com.client as client
import datetime
import datacompy

ctx1 = snowflake.connector.connect(
          user='MOURADABIDI',
          password='Ma@07842032',
          account='ba62849.east-us-2.azure',
          warehouse= 'COMPUTE_MACHINE',
          database1='DB_ASEA_REPORTS',
          schema1='DBO'
        #   schema='SNAPSHOT'
          )    

ctx2 = snowflake.connector.connect(
          user='MOURADABIDI',
          password='Ma@07842032',
          account='ba62849.east-us-2.azure',
          warehouse= 'COMPUTE_MACHINE',
          database2='DB_RAW_DATA',
          schema2='INFOTRAX_PROD')


def  SnowflakeQA(Table1, Table2):
# def  SnowflakeQA(Table):    
    cur1 = ctx1.cursor()
    
# # Execute a statement that will generate a result set.
    warehouse= 'COMPUTE_MACHINE'
    database1='DB_ASEA_REPORTS'
    schema1='DBO'
   
    if warehouse:
        cur1.execute(f'use warehouse {warehouse};')
    # cur1.execute(f'select * from {database1}.{schema1}.{Table1};')

    # cur1.execute("""SELECT DISTRIBUTORID
    # ,RANK_CHANGE
    # , CAST(STARTDATE AS DATE) AS STARTDATE
    # , CAST(ENDDATE AS DATE) AS ENDDATE
    # , ACCOUNTTYPE
    # , POST_LOG_UPDATE
    # FROM DB_RAW_DATA.INFOTRAX_PROD.TBL_DISTRIBUTOR;""")

    cur1.execute(f'select * from {database1}.{schema1}.{Table1};')
    #cur.execute("""SELECT * FROM DB_ASEA_REPORTS.PUBLIC.INFOPROD_RANKS WHERE EnteredDate < '2021-07-21'""")
# Fetch the result set from the cursor and deliver it as the Pandas DataFrame.
    snowflakedf1 = cur1.fetch_pandas_all()
    
    
    cur2 = ctx2.cursor()
    warehouse= 'COMPUTE_MACHINE'
    database2='DB_RAW_DATA'
    schema2='INFOTRAX_PROD'


    if warehouse:
        cur2.execute(f'use warehouse {warehouse};')
    cur2.execute(f'select * from {database2}.{schema2}.{Table2};')

    # cur2.execute("""SELECT DISTRIBUTORID
    # ,RANK_CHANGE
    # , CAST(STARTDATE AS DATE) AS STARTDATE
    # , CAST(ENDDATE AS DATE) AS ENDDATE
    # , ACCOUNTTYPE
    # , POST_LOG_UPDATE
    # FROM DB_ASEA_REPORTS.STAGE.TBL_DISTRIBUTOR_TEMP;""")


    cur2.execute(f'select * from {database2}.{schema2}.{Table2};')
    snowflakedf2 = cur2.fetch_pandas_all()





    # snowflakedf1 = snowflakedf[(snowflakedf.CREATEDDATE > '2018-01-01' ) & (snowflakedf.CREATEDDATE < start_date) & (snowflakedf.UPDATEDDATE <= start_date)] 
    # # snowflakeddf1 = snowflakedf.query('CREATEDDATE > "2015-01-10"
    # #            and CREATEDDATE < start_date 
    # #            and UPDATEDDATE <= start_date', inplace = True)

    compare = datacompy.Compare(
    snowflakedf1,
    snowflakedf2,
    

    #join_columns= 'LEGACYNUMBER')
    join_columns= ['LEGACYNUMBER'])
    #join_columns= Primary_key)


    compare.matches(ignore_extra_columns=False) 
    print(compare.report())
    #sqldatabase = 'InfoTrax_Prod'
    # sqldatabase = 'ASEA_PROD'
    # sqldatabase = 'ASEA_REPORTS'
    sqldatabase ='DB_ASEA_REPORTS'
    Today = datetime.datetime.today()
    outlook = client.Dispatch('Outlook.Application')
    message = outlook.Createitem(0)
    message.Display()
    message.To = 'mabidi@aseaglobal.com'
    message.Subject = 'SQL COMPARE With Window Time '  + sqldatabase +'.dbo.' + ' ' + ' as of ' + ' ' + str(Today)
    message.Body = compare.report()
    message.Save()
    message.Send()

    cur1.close()
    cur2.close()

result = SnowflakeQA('TBL_DISTRIBUTOR', 'TBL_DISTRIBUTOR')