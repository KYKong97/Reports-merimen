from datetime import datetime, timedelta
import pandas as pd
import logging
import os
import win32com.client


def write_to_macro(macro_excel:str, df:pd.DataFrame, output_file:str=None):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(macro_excel)
    ws = wb.ActiveSheet
    
    

    no_column = df['No'].tolist()
    claim_no_column = df['Claim No'].tolist()
    liable_amount_column = df['Solicitor Worksheet Liable Amount'].tolist()

    for i,cell in enumerate(no_column):
        ws.Cells(6 + i,1).value = cell

    for i,cell in enumerate(claim_no_column):
        ws.Cells(6 + i, 2).value = cell

    for i,cell in enumerate(liable_amount_column):
        ws.Cells(6 + i,7).value = cell

    if output_file is None:
        wb.SaveAs(macro_excel)
    else:
        wb.SaveAs(output_file)
    
    excel.Application.Quit()


def run_macro(macro_excel):
    xl=win32com.client.Dispatch("Excel.Application")
    xl.Workbooks.Open(macro_excel, ReadOnly=1)
    basename = os.path.basename(macro_excel)
    xl.Application.Run(f"'{basename}'!Button26_DATAPREPARATION")
##    xl.Application.Save() # if you want to save then uncomment this line and change delete the ", ReadOnly=1" part from the open function.
    xl.Application.Quit() # Comment this out if your excel script closes

def preprocess_df(df:pd.DataFrame):
    try:
        df['Solicitor Worksheet Liable Amount'] = df['Solicitor Worksheet Liable Amount'].str.replace('[^\d.]', '', regex=True)
        df['Solicitor Worksheet Liable Amount'] = pd.to_numeric(df['Solicitor Worksheet Liable Amount'], errors='coerce')
        df['Claim No'] = df['Claim No'].str.replace(r'-.+', '', regex=True)
        grouped_df = df.groupby(['Claim No', 'Latest Solicitor Opinion Report Submitted Date']).agg({
            'Solicitor Worksheet Liable Amount': 'sum',
            'No': 'first',
            'Panel Solicitor': 'first',
            'PIC': 'first',
            'Panel Solicitor Assigned Date': 'first',
        }).reset_index()

        column_order = ['No', 'Claim No', 'Solicitor Worksheet Liable Amount', 'Panel Solicitor', 'PIC', 'Panel Solicitor Assigned Date', 'Latest Solicitor Opinion Report Submitted Date']
        grouped_df = grouped_df[column_order]
        grouped_df['No'] = range(1, len(grouped_df) + 1)
        return grouped_df
    except Exception as e:
        logging.error(f"Error in preprocess df Error msg:{e}")
        return -1



def date_range(option_selected:str):
    
    if option_selected == "Last Year":
        today = datetime.now()
        first_day_of_last_year = today.replace(year=today.year - 1, month=1, day=1)
        last_day_of_last_year = today.replace(year=today.year - 1, month=12, day=31)
        from_date_input = first_day_of_last_year.strftime("%d/%m/%Y")
        to_date_input = last_day_of_last_year.strftime("%d/%m/%Y")
    elif option_selected == "Current Year":
        today = datetime.now()
        first_day_of_current_year = today.replace(month=1, day=1)
        last_day_of_current_year = today.replace(month=12, day=31)
        from_date_input = first_day_of_current_year.strftime("%d/%m/%Y")
        to_date_input = last_day_of_current_year.strftime("%d/%m/%Y")
    elif option_selected == "Last Month":
        today = datetime.now()
        first_day_of_current_month = today.replace(day=1)
        last_day_of_last_month = first_day_of_current_month - timedelta(days=1)
        last_day_of_last_month = last_day_of_last_month.replace(day=1)
        from_date_input = last_day_of_last_month.strftime("%d/%m/%Y")
        to_date_input = (first_day_of_current_month - timedelta(days=1)).strftime("%d/%m/%Y")
    elif option_selected == "Current Month":
        today = datetime.now()
        first_day_of_current_month = today.replace(day=1)
        last_day_of_current_month = today
        from_date_input = first_day_of_current_month.strftime("%d/%m/%Y")
        to_date_input = last_day_of_current_month.strftime("%d/%m/%Y")
    elif option_selected == "Last Week":
        today = datetime.now()
        last_week = today - timedelta(weeks=1)
        first_day_of_last_week = last_week - timedelta(days=last_week.weekday())
        from_date_input = first_day_of_last_week.strftime("%d/%m/%Y")
        to_date_input = last_week.strftime("%d/%m/%Y")
    elif option_selected == "Current Week":
        today = datetime.now()
        first_day_of_current_week = today - timedelta(days=today.weekday())
        from_date_input = first_day_of_current_week.strftime("%d/%m/%Y")
        to_date_input = today.strftime("%d/%m/%Y")
    elif option_selected == "Yesterday":
        today = datetime.now()
        yesterday = today - timedelta(days=1)
        from_date_input = yesterday.strftime("%d/%m/%Y")
        to_date_input = yesterday.strftime("%d/%m/%Y")
    elif option_selected == "Today":
        today = datetime.now()
        from_date_input = today.strftime("%d/%m/%Y")
        to_date_input = today.strftime("%d/%m/%Y")
    
    return from_date_input, to_date_input
