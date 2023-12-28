from lib import generate_report_after_macro
import win32com.client

xl=win32com.client.Dispatch("Excel.Application")
xl.Application.DisplayAlerts = False
wb = xl.Workbooks.Open(r"D:\Projects\TPBI_Report\data\TPBI UPDATE RESERVE MACRO_output_20231227.xlsm")
from_date="test"
to_date = "test"
today="test"

generate_report_after_macro(xl, wb, from_date,to_date,today)


