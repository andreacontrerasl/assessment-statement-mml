import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.table import Table, TableStyleInfo

def create_excel_report(df):
    path_to_save = 'Financial_Analysis.xlsx'
    
    with pd.ExcelWriter(path_to_save, engine='openpyxl') as writer:
        pivot_df = df.pivot_table(index=['Client', 'Currency'], values='Transaction', aggfunc='count')
        pivot_df.to_excel(writer, sheet_name='Transactions Summary')
        
        totals_usd = df.groupby('Client')['Transaction'].sum()
        totals_usd.to_excel(writer, sheet_name='Total by Client USD')

        workbook = writer.book
        sheet1 = workbook['Transactions Summary']
        sheet2 = workbook['Total by Client USD']
        
        # Set AutoFilter on the DataFrame ranges
        sheet1.auto_filter.ref = sheet1.dimensions
        sheet2.auto_filter.ref = sheet2.dimensions
        
        chart = BarChart()
        values = Reference(sheet2, min_col=2, min_row=2, max_row=sheet2.max_row)
        categories = Reference(sheet2, min_col=1, min_row=2, max_row=sheet2.max_row)
        chart.add_data(values, titles_from_data=True)
        chart.set_categories(categories)
        chart.title = "Total USD per Client"
        sheet2.add_chart(chart, "E4")
        
        tab = Table(displayName="TransactionTable", ref="A1:C{}".format(pivot_df.shape[0]+1))
        style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                               showLastColumn=False, showRowStripes=True, showColumnStripes=True)
        tab.tableStyleInfo = style
        sheet1.add_table(tab)

def create_excel_report_graph(df):
    path_to_save = 'Financial_Analysis.xlsx'
    
    # Clean data
    df.dropna(subset=['Client', 'Currency'], inplace=True)  # Ensure no NaN values in these columns
    
    # Create a Pandas Excel writer using Openpyxl
    with pd.ExcelWriter(path_to_save, engine='openpyxl') as writer:
        # Total transactions per client, per currency
        pivot_df = df.pivot_table(index=['Client', 'Currency'], values='Transaction', aggfunc='count')
        pivot_df.to_excel(writer, sheet_name='Transactions Summary')
        
        # Totals by client in USD
        totals_usd = df.groupby('Client')['Transaction'].sum()
        totals_usd.to_excel(writer, sheet_name='Total by Client USD')

        workbook = writer.book
        sheet1 = workbook['Transactions Summary']
        sheet2 = workbook['Total by Client USD']
        
        # Set AutoFilter on the DataFrame ranges
        sheet1.auto_filter.ref = sheet1.dimensions
        sheet2.auto_filter.ref = sheet2.dimensions
    
    unique_clients = df['Client'].dropna().unique() 
    
    for client in unique_clients:
        # Filter data for the current client
        client_data = pivot_df.xs(client, level='Client')
        client_data.reset_index(inplace=True)  # Reset index to ensure integer indexing

        try:
            min_row_index = int(client_data.index[0]) + 2
            max_row_index = int(client_data.index[-1]) + 2
        except ValueError as ve:
            print(f"Error converting index to integer: {ve}")
            continue

        data = Reference(sheet1, min_col=2, min_row=min_row_index, max_row=max_row_index)
        cats = Reference(sheet1, min_col=1, min_row=min_row_index, max_row=max_row_index)
        
        chart = BarChart()
        chart.title = f"Transactions per Currency for {client}"
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        sheet1.add_chart(chart, f"A{max_row_index + 4}")

    
    # Create a bar chart for total by client USD
    chart2 = BarChart()
    values = Reference(sheet2, min_col=2, min_row=2, max_row=sheet2.max_row)
    categories = Reference(sheet2, min_col=1, min_row=2, max_row=sheet2.max_row)
    chart2.add_data(values, titles_from_data=True)
    chart2.set_categories(categories)
    chart2.title = "Total USD per Client"
    sheet2.add_chart(chart2, "E4")
        
    # Add formatting and a table style to the pivot table sheet
    tab = Table(displayName="TransactionTable", ref="A1:C{}".format(pivot_df.shape[0]+1))
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                           showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    tab.tableStyleInfo = style
    sheet1.add_table(tab)

    # Save the workbook
    workbook.save(filename=path_to_save)