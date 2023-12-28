from src.merimen_controller import MerimenController
from src.lib import date_range
from datetime import datetime, timedelta
from src.lib import preprocess_df, write_to_macro, run_macro, generate_lf_fees,generate_report_after_macro
import tkinter  as tk 
import sys
from tkinter import messagebox
from tkcalendar import DateEntry
import os
TODAY = datetime.today().strftime("%d/%m/%Y")
TODAY_file = datetime.today().strftime("%Y%m%d")

def warm_up_merimen(merimen_username,merimen_password,from_date, to_date):
    if from_date is None:
        from_date = datetime.today().strftime("%d/%m/%Y")
    if to_date is None:
        to_date = datetime.today().strftime("%d/%m/%Y")
    if isinstance(from_date, datetime):
        from_date = from_date.strftime("%d/%m/%Y")
    if isinstance(to_date,datetime):
        to_date = to_date.strftime("%d/%m/%Y")

    if merimen_username is None or len(merimen_username)==0:
        messagebox.showerror("Please enter merimen username")
        return 
    if merimen_password is None or len(merimen_password)==0:
        messagebox.showerror("Please enter merimen password")
        return
    
    merimen_controller = MerimenController(merimen_username=merimen_username,merimen_password=merimen_password)
    start_merimen(merimen_controller,from_date, to_date)
    merimen_controller.exit()
    

def start_merimen(merimen_controller,from_date, to_date):
    
    

    
    merimen_controller.filter_claim_type("TP BI")
    try:
        merimen_controller.merimen_page.get_by_role("link", name="Report on Outstanding TPBI Claims").click(timeout=1000000)
    except Exception as e:
        messagebox.showerror("Cannot find Report on Outstanding TPBI Claims")
        return
    

    merimen_controller.filter_report_date(from_date=from_date,to_date=to_date)
    merimen_controller.check_opinion_report()
    merimen_controller.generate_report()
    report_df = merimen_controller.read_report_table()
    if len(report_df)==0:
        messagebox.showinfo("Empty","Empty table in merimen")
        return
    
    df_output = preprocess_df(report_df)
    df_output = generate_lf_fees(df_output)

    data_folder = os.path.join(os.getcwd(),"data")
    macro_excel = os.path.join(data_folder,"TPBI UPDATE RESERVE MACRO.xlsm")
    output_macro_excel = os.path.join(data_folder,f"TPBI UPDATE RESERVE MACRO_output_{TODAY_file}.xlsm")
    
    try:
        write_to_macro(macro_excel, df_output, output_file=output_macro_excel)
    except Exception as e:
        messagebox.showerror("Cannot write to macro",str(e))
        return

    messagebox.showinfo("Done","Click ok to continue for macro")

    try:
        xl,wb = run_macro(output_macro_excel)
        generate_report_after_macro(xl, wb, from_date,to_date,TODAY)
        messagebox.showinfo("Run macro done","Run Macro Done")
    except Exception as e:
        messagebox.showerror("Cannot generate macro record",f"Cannot run macro due to {e}")
        return

    
    # merimen_controller.merimen_page.pause()
    
class MainWindow:
    def __init__(self) -> None:
        self.final_date_from = None
        self.final_date_to = None
        self.my_w = tk.Tk()
        self.my_w.geometry("340x220")  

        self.merimen_username_label = tk.Label(self.my_w,text="Merimen Username")
        self.merimen_username_label.grid(row=1,column=1)
        self.merimen_username_input = tk.Entry(self.my_w)
        self.merimen_username_input.grid(row=1,column=2)

        self.merimen_password_label = tk.Label(self.my_w,text="Merimen Password")
        self.merimen_password_label.grid(row=2,column=1)
        self.merimen_password_input = tk.Entry(self.my_w)
        self.merimen_password_input.grid(row=2,column=2)



            ## Label
        self.from_label = tk.Label(self.my_w,text="From Date")
        self.from_label.grid(row=3,column=1)

        self.to_label = tk.Label(self.my_w,text="To Date")
        self.to_label.grid(row=4,column=1)

        self.from_cal = DateEntry(self.my_w, selectmode="day", date_pattern="dd-MM-yyyy")
        self.from_cal.grid(row=3,column=2, padx=15)

        self.to_cal = DateEntry(self.my_w,selectmode="day", date_pattern="dd-MM-yyyy")
        self.to_cal.grid(row=4, column=2, padx=15)

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
        self.space_label.grid(row=5,column=1)
        self.question_menu.grid(row=6,column=2)

        self.result_label = tk.Label(self.my_w)
        self.result_label.grid(row=7,column=1)

        self.to_cal.bind("<<DateEntrySelected>>", self.update_label)  
        self.from_cal.bind("<<DateEntrySelected>>", self.update_label)  

        self.sub_btn = tk.Button(self.my_w,text="Submit",command=lambda :warm_up_merimen(self.merimen_username_input.get(),self.merimen_password_input.get(),self.final_date_from,self.final_date_to))

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
    # import win32com
    # xl=win32com.client.Dispatch("Excel.Application")
    # wb = xl.Workbooks.Open(r"D:\Projects\TPBI_Report\data\TPBI UPDATE RESERVE MACRO_output_20231218.xlsm", ReadOnly=1)

    # generate_report_after_macro(xl, wb, "20/12/2023","20/12/2023","19/12/2023")

    # if len(sys.argv)>2:
    #     data_folder = os.path.join(os.getcwd(),"data")
    #     output_macro_excel = os.path.join(data_folder,"TPBI UPDATE RESERVE MACRO_output.xlsm")
    #     run_macro(output_macro_excel)
    # else:
    main = MainWindow()
    