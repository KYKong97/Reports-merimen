from datetime import datetime, timedelta
import pandas as pd
import logging
import os
import win32com.client
import json
import math

with open("data/config.json","r") as f:
        east_malaysia = json.load(f)['east']


def lf_fees_mapping(settlement_amount,panel_solicitor, nearest=500.0):
    reserve_amount = None
    if panel_solicitor in east_malaysia:
        if 1<= settlement_amount <= 20000:
            reserve_amount = 3000
        elif 20001<=settlement_amount <=100000:
            reserve_amount = round(0.09*settlement_amount)
        elif 100001<=settlement_amount<=500000:
            reserve_amount = round(0.08*settlement_amount)
        elif settlement_amount>500001:
            reserve_amount = round(0.05*settlement_amount)
            reserve_amount = 30000 if reserve_amount>30000 else reserve_amount
    else:
        if 1<= settlement_amount <= 20000:
            reserve_amount = 3000
        elif 20001<=settlement_amount <=50000:
            reserve_amount = round(0.08*settlement_amount)
        elif 50001<=settlement_amount<=100000:
            reserve_amount = round(0.07*settlement_amount)
        elif 100001<=settlement_amount<=500000:
            reserve_amount = round(0.05*settlement_amount)
        elif settlement_amount>500001:
            reserve_amount = round(0.05*settlement_amount)
            reserve_amount = 30000 if reserve_amount>30000 else reserve_amount
    
    ## Round down to nearest 500
    reserve_amount = math.ceil(reserve_amount/nearest)*nearest
    reserve_amount = int(reserve_amount)
    return reserve_amount


def generate_lf_fees(df:pd.DataFrame):
    df['Solicitor_Fees'] = df.apply(lambda x:lf_fees_mapping(x['Solicitor Worksheet Liable Amount'],x['Panel Solicitor']),axis=1)
    return df

    

def write_to_macro(macro_excel:str, df:pd.DataFrame, output_file:str=None):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.EnableEvents = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(macro_excel)
    # ws = wb.ActiveSheet
    ws = wb.Worksheets("claims")
    
    

    no_column = df['No'].tolist()
    claim_no_column = df['Claim No'].tolist()
    liable_amount_column = df['Solicitor Worksheet Liable Amount'].tolist()
    solicitor_fees = df['Solicitor_Fees'].tolist()

    

    for i,cell in enumerate(no_column):
        ws.Cells(6 + i,1).value = cell

    for i,cell in enumerate(claim_no_column):
        ws.Cells(6 + i, 2).value = cell

    for i,cell in enumerate(liable_amount_column):
        ws.Cells(6 + i,7).value = cell
    
    for i, cell in enumerate(solicitor_fees):
        ws.Cells(6+i,8).value = cell


    if output_file is None:        
        wb.SaveAs(macro_excel,ConflictResolution=2)
    else:
        

        wb.SaveAs(output_file,ConflictResolution=2)
    
    # excel.Application.Quit()


def run_macro(macro_excel):
    xl=win32com.client.Dispatch("Excel.Application")
    xl.Application.DisplayAlerts = False
    wb = xl.Workbooks.Open(macro_excel)
    
    basename = os.path.basename(macro_excel)
    xl.Application.Run(f"'{basename}'!Button26_DATAPREPARATION")
    xl.Application.Run(f"'{basename}'!Button1_Inquiry")
    wb.SaveAs(macro_excel,ConflictResolution=2)
    # xl.Application.Quit() # Comment this out if your excel script closes
    return xl,wb

def generate_individual_report(report_worksheet, list_claims, today,from_date, to_date):
    report_worksheet.Cells(2,3).value = today
    report_worksheet.Cells(3,3).value = f"{from_date}-{to_date}"
    report_first_row = 8
    sum_row = report_first_row+len(list_claims)
    for index,data in enumerate(list_claims):
        report_worksheet.Cells(report_first_row+index,2).value = data['No']
        report_worksheet.Cells(report_first_row+index,3).value = data['Claim No']
        report_worksheet.Cells(report_first_row+index,4).value = data['Before_BI_Reserve']
        report_worksheet.Cells(report_first_row+index,5).value = data['Before_LF_Reserve']
        report_worksheet.Cells(report_first_row+index,6).value = data['After_BI_Reserve']
        report_worksheet.Cells(report_first_row+index,7).value = data['After_LF_Reserve']
        report_worksheet.Cells(report_first_row+index,8).value = data['Movement_BI']
        report_worksheet.Cells(report_first_row+index,9).value = data['Movement_LF']
        report_worksheet.Cells(report_first_row+index,10).value = data['Total_Movement']
    
    report_worksheet.Cells(sum_row,3).value = "Total"
    report_worksheet.Cells(sum_row,4).value = f"=SUM(D{report_first_row}:D{sum_row-1})"
    report_worksheet.Cells(sum_row,5).value = f"=SUM(E{report_first_row}:E{sum_row-1})"
    report_worksheet.Cells(sum_row,6).value = f"=SUM(F{report_first_row}:F{sum_row-1})"
    report_worksheet.Cells(sum_row,7).value = f"=SUM(G{report_first_row}:G{sum_row-1})"
    report_worksheet.Cells(sum_row,8).value = f"=SUM(H{report_first_row}:H{sum_row-1})"
    report_worksheet.Cells(sum_row,9).value = f"=SUM(I{report_first_row}:I{sum_row-1})"
    report_worksheet.Cells(sum_row,10).value = f"=SUM(J{report_first_row}:J{sum_row-1})"


def generate_report_after_macro(excel, workbook, from_date,to_date,today):
    claims_worksheet = workbook.Worksheets("claims")
    egib_report_worksheet = workbook.Worksheets("EGIB_Report")
    egtb_report_worksheet = workbook.Worksheets("EGTB_Report")

    lastRow = claims_worksheet.UsedRange.Rows.Count
    list_egib_claims = []
    list_egtb_claims = []
    first_row = 6
    for i in range(lastRow):
        no = claims_worksheet.Cells(first_row+i,1).value
        if no is None:
            break
        claim_no = claims_worksheet.Cells(first_row+i,2).value
        new_bi_reserve = claims_worksheet.Cells(first_row+i,7).value
        new_lf_reserve = claims_worksheet.Cells(first_row+i,8).value
        reserve_bi_polm = claims_worksheet.Cells(first_row+i,11).value
        if reserve_bi_polm is None:
            reserve_bi_polm = 0
        reserve_lf_polm = claims_worksheet.Cells(first_row+i,12).value
        if reserve_lf_polm is None:
            reserve_lf_polm = 0

        temp_dict = {
            "No": no,
            "Claim No":claim_no,
            "Before_BI_Reserve":reserve_bi_polm,
            "Before_LF_Reserve":reserve_lf_polm,
            "After_BI_Reserve":new_bi_reserve,
            "After_LF_Reserve":new_lf_reserve,
            "Movement_BI":abs(new_bi_reserve-reserve_bi_polm),
            "Movement_LF":abs(new_lf_reserve - reserve_lf_polm),
            "Total_Movement":abs(new_bi_reserve-reserve_bi_polm)+abs(new_lf_reserve - reserve_lf_polm)
        }
        if claim_no[0].upper()=="K":
            list_egtb_claims.append(temp_dict)
        elif claim_no[0].upper()=="V":
            list_egib_claims.append(temp_dict)
        else:
            raise Exception(f"Claim No {claim_no} is not start with K or V")

    
    generate_individual_report(egib_report_worksheet, list_egib_claims, today,from_date, to_date)
    generate_individual_report(egtb_report_worksheet, list_egtb_claims, today,from_date, to_date)

    workbook.Close(SaveChanges=1)





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
        grouped_df = grouped_df.round(2)
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
