import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
from datetime import datetime


def create_deliverable_2(df):
    path_to_save = 'Latin_America_Transactions.xlsx'
    
    df['Date'] = pd.to_datetime(df['Date']).dt.date
    
    # Filtering transactions for Latin American countries
    latin_american_countries = [
    'ARGENTINA', 'BOLIVIA', 'BRAZIL', 'CHILE', 'COLOMBIA', 
    'COSTA RICA', 'CUBA', 'DOMINICAN REPUBLIC', 'ECUADOR', 
    'EL SALVADOR', 'GUATEMALA', 'HONDURAS', 'MEXICO', 
    'NICARAGUA', 'PANAMA', 'PARAGUAY', 'PERU', 
    'PUERTO RICO', 'URUGUAY', 'VENEZUELA'
    ]

    df_latam = df[df['Country'].isin(latin_american_countries)]

    specific_countries = ['RUSSIA', 'CUBA', 'CHINA', 'VENEZUELA']
    
    start_date = datetime.strptime('2024-10-01', '%Y-%m-%d').date()
    end_date = datetime.strptime('2024-12-31', '%Y-%m-%d').date()
    
    df_specific = df[(df['Country'].isin(specific_countries)) & 
                     (df['Date'] >= start_date) & 
                     (df['Date'] <= end_date)]

    cols_basic = ['Client', 'Country', 'Currency', 'Transaction']
    cols_with_date = cols_basic + ['Date']
    
    with pd.ExcelWriter(path_to_save, engine='openpyxl') as writer:
        df_latam.loc[:, cols_basic].to_excel(writer, sheet_name='LATAM Transactions')
        df_specific.loc[:, cols_with_date].to_excel(writer, sheet_name='Specified Countries Q4 2024')
        
        workbook = writer.book
        sheet1 = workbook['LATAM Transactions']
        sheet2 = workbook['Specified Countries Q4 2024']
        
        sheet1.auto_filter.ref = sheet1.dimensions
        sheet2.auto_filter.ref = sheet2.dimensions
        
    workbook.save(filename=path_to_save)