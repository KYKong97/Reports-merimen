from src.merimen_controller import MerimenController
from src.lib import date_range, process_and_write_data
from datetime import datetime, timedelta
from src.lib import preprocess_df, write_to_macro, run_macro
import tkinter  as tk 
from tkcalendar import DateEntry
import os
import pandas as pd
import numpy as np

def start_merimen(from_date, to_date):

    if isinstance(from_date, datetime):
        from_date = from_date.strftime("%d/%m/%Y")
    if isinstance(to_date,datetime):
        to_date = to_date.strftime("%d/%m/%Y")

    merimen_controller = MerimenController(merimen_username="MIGAMINUR",merimen_password="R10.5.2004n")
    merimen_controller.filter_claim_type("TP BI")
    merimen_controller.merimen_page.get_by_role("link", name="Report on Outstanding TPBI Claims").click(timeout=1000000)

    merimen_controller.filter_report_date(from_date=from_date,to_date=to_date)
    merimen_controller.check_opinion_report()
    merimen_controller.generate_report()
    report_df = merimen_controller.read_report_table()
    df_output = preprocess_df(report_df)

    macro_excel = os.path.join(os.getcwd(),"TPBI UPDATE RESERVE MACRO.xlsm")
    output_macro_excel = os.path.join(os.getcwd(),"TPBI UPDATE RESERVE MACRO_output.xlsm")
    write_to_macro(macro_excel, df_output, output_file=output_macro_excel)
    
#Change Kong code start here 
    process_and_write_data(df, 'C:\\Users\\acer\\Desktop\\kerja\\Code optimized\\TPBI UPDATE RESERVE MACRO.xlsm', 'C:\\Users\\acer\\Desktop\\kerja\\Code optimized\\TPBI UPDATE RESERVE MACRO_output.xlsm')
#Change Kong code end here    

    run_macro(output_macro_excel)
    merimen_controller.merimen_page.pause()
    
   
   
    
class MainWindow:
    def __init__(self) -> None:
        self.final_date_from = None
        self.final_date_to = None
        self.my_w = tk.Tk()
        self.my_w.geometry("340x220")  
            ## Label
        self.from_label = tk.Label(self.my_w,text="From Date")
        self.from_label.grid(row=1,column=1)

        self.to_label = tk.Label(self.my_w,text="To Date")
        self.to_label.grid(row=3,column=1)

        self.from_cal = DateEntry(self.my_w, selectmode="day", date_pattern="dd-MM-yyyy")
        self.from_cal.grid(row=1,column=2, padx=15)

        self.to_cal = DateEntry(self.my_w,selectmode="day", date_pattern="dd-MM-yyyy")
        self.to_cal.grid(row=3, column=2, padx=15)

        options_list = ["Last Year", "Current Year", "Last Month", "Current Month",
                        "Last Week","Current Week","Yesterday","Today"] 
    
        # Variable to keep track of the option 
        # selected in OptionMenu 
        self.value_inside = tk.StringVar(self.my_w) 
        
        # Set the default value of the variable 
        self.value_inside.set("Select an Option") 
        
        # Create the optionmenu widget and passing  
        # the options_list and value_inside to it.

        self.question_menu = tk.OptionMenu(self.my_w, self.value_inside, *options_list, command=self.update_label_from_option) 
        self.space_label = tk.Label(self.my_w,text="")
        self.space_label.grid(row=4,column=1)
        self.question_menu.grid(row=5,column=2)

        self.result_label = tk.Label(self.my_w)
        self.result_label.grid(row=7,column=1)

        self.to_cal.bind("<<DateEntrySelected>>", self.update_label)  
        self.from_cal.bind("<<DateEntrySelected>>", self.update_label)  

        self.sub_btn = tk.Button(self.my_w,text="Submit",command=lambda :start_merimen(self.final_date_from,self.final_date_to))

        self.sub_btn.grid(row=10,column=2)
        self.my_w.mainloop()
        

    def update_label_from_option(self,selection):
        final_date_from, final_date_to = date_range(selection)
        self.from_cal.set_date(datetime.strptime(final_date_from,"%d/%m/%Y"))
        self.to_cal.set_date(datetime.strptime(final_date_to,"%d/%m/%Y"))
        self.result_label['text'] = "From {} to {}".format(final_date_from, final_date_to)
        
        self.final_date_from = datetime.strptime(final_date_from,"%d/%m/%Y")
        self.final_date_to = datetime.strptime(final_date_to,"%d/%m/%Y")

    def update_label(self,event):
        final_date_from = self.from_cal.get_date().strftime("%d-%m-%Y")
        final_date_to = self.to_cal.get_date().strftime("%d-%m-%Y")
        self.result_label['text'] = "From {} to {}".format(final_date_from, final_date_to)

        self.final_date_to = final_date_to
        self.final_date_from = final_date_from

 
    

    
    
    

    

if __name__=="__main__":
    main = MainWindow()
    