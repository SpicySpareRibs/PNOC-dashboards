import pandas
import tkinter as tk
from tkinter import ttk, font, filedialog as fd
from typing import get_args

from annual_procurement_plan_recompiler import AnnualProcurementPlanRecompiler
from recompiler import Datafile, filetype
from plantilla_recompiler import PlantillaOfPersonnelCompiler
from training_summary_recompiler import TrainingSummaryRecompiler
from contract_recompiler import ContractCompiler
from proc_monitorning_report import ProcMonitoringCompiler
from inventory_of_clients_recompiler import InventoryOfClientsRecompiler
from staffing_summary_recompiler import StaffingSummaryRecompiler

class RecompilerMaker():
    def make(datafile: Datafile):
        match datafile.filetype:
            case "Training Summary Report":
                return TrainingSummaryRecompiler(datafile)
            case "Plantilla of Personnel":
                return PlantillaOfPersonnelCompiler(datafile)
            case "Contract of Service":
                return ContractCompiler(datafile)
            case "Procurement Monitoring Report":
                return ProcMonitoringCompiler(datafile)
            case "Inventory of Clients":
                return InventoryOfClientsRecompiler(datafile)
            case "Staffing Summary":
                return StaffingSummaryRecompiler(datafile)
            case "Annual Procurement Plan":
                return AnnualProcurementPlanRecompiler(datafile)
            case _:
                print("Invalid Input")
                return None

class GUI():
    def __init__(self):
        # Initialize the window
        self.root = tk.Tk()
        self.root.title("PNOC Datafile Recompiler")
        self.root.geometry("750x600")
        self.root.resizable(0,0)

        self.default_font = font.nametofont("TkDefaultFont")
        self.default_font.configure(family="Segoe UI", size=10, weight="normal")
        self.text_font = font.nametofont("TkTextFont")
        self.text_font.configure(family="Segoe UI", size=10, weight="normal")
        self.bold_text_font = ["Segoe UI", 10, "bold"]

        self.selected_datafile = tk.StringVar()

        # Initialize widgets
        header_bg = tk.Canvas(self.root, width=750, height=115, bg='#991E20')
        header_bg.create_line(0, 116, 750, 116, fill='#A6A6A6', width=2)
        header_bg.place(x=0, y=0)
        
        footer = tk.Canvas(self.root, width=750, height=5)
        footer.create_line(0,0,750,0, fill='#A6A6A6', width=5)
        footer.place(x=160, y=515)

        footer_text = tk.Label(self.root, font=["Segoe UI", 8, "normal"], foreground='#A6A6A6')
        footer_text['text'] = "PNOC Datafile Recompiler"
        footer_text.place(x=5, y=505)

        title_text = ttk.Label(self.root, foreground='#FFF9E9', font=["Segoe UI", 10, "bold"], background='#991E20')
        title_text['text'] = "Information"
        title_text.place(x=40, y=20)

        information_text = ttk.Label(self.root, wraplength=650, foreground='#FFF9E9', background='#991E20')
        information_text['text'] = "This application recompiles the selected datafile into a Looker Studio compatible format. The output file can be located in the same folder as the application."
        information_text.place(x=60, y=45)

        select_recompiler_label = ttk.Label(self.root, font=self.bold_text_font)
        select_recompiler_label['text'] = "Select Recompiler Type:"
        select_recompiler_label.place(x=60, y=140)

        self.selected_recompiler = tk.StringVar()
        select_recompiler_combobox = ttk.Combobox(self.root, width = 27, textvariable = self.selected_recompiler, postcommand = self.reset_status, state='readonly')
        select_recompiler_combobox['values'] = get_args(filetype)
        select_recompiler_combobox.place(x=270, y=140)
        select_recompiler_combobox.current(0)

        recompiler_information = ttk.Label(self.root, wraplength=630)
        recompiler_information['text'] = "The recompiler type is the type of file to be recompiled. The dropdown menu contains the supported types of files to be recompiled."
        recompiler_information.place(x=80, y=175)

        select_datafile_label = ttk.Label(self.root, font=self.bold_text_font)
        select_datafile_label['text'] = "Select the Input Datafile"
        select_datafile_label.place(x=60, y=250)

        select_datafile_entry = ttk.Entry(self.root, width=53, textvariable=self.selected_datafile)
        select_datafile_entry.place(x=80, y=285)

        select_datafile_button = ttk.Button(self.root, text="Browse...", command=self.get_datafile)
        select_datafile_button.place(x=580, y=283)

        select_datafile_information = ttk.Label(self.root, wraplength=630)
        select_datafile_information['text'] = "Locate on your device the specific datafile to be recompiled."
        select_datafile_information.place(x=80, y=325)

        select_sheet_label = ttk.Label(self.root, font=self.bold_text_font)
        select_sheet_label['text'] = "Select Sheet: "
        select_sheet_label.place(x=60, y=375)

        self.selected_sheet = tk.StringVar()
        self.select_sheet_combobox = ttk.Combobox(self.root, width = 27, textvariable = self.selected_sheet, postcommand = self.reset_status, state='readonly')
        self.select_sheet_combobox['values'] = []
        self.select_sheet_combobox.place(x=180, y=375)

        select_sheet_information = ttk.Label(self.root, wraplength=630)
        select_sheet_information['text'] = "Select the specific sheet of the datafile to be recompiled."
        select_sheet_information.place(x=80, y=410)

        recompile_button = ttk.Button(self.root, text="Recompile", command=self.recompile)
        recompile_button.place(x=580, y=540)

        self.recompile_status = ttk.Label(self.root)
        self.recompile_status['text'] = ""
        self.recompile_status.place(x=420, y=542)

        self.root.mainloop()
    
    def reset_status(self) -> None:
        self.recompile_status['text'] = ""

    def get_datafile(self) -> None:
<<<<<<< HEAD
        accepted_filetypes = [("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        self.selected_datafile.set(fd.askopenfilename(filetypes=accepted_filetypes))
        sheets = list(pandas.read_excel(self.selected_datafile.get(), sheet_name=None).keys())
        self.select_sheet_combobox['values'] = sheets
        self.select_sheet_combobox.current(0)
=======
        self.selected_datafile.set(fd.askopenfilename())
>>>>>>> 7966cef (Fixed bug on selecting datafiles)
        self.reset_status()

    def recompile(self) -> None:
<<<<<<< HEAD
        datafile = Datafile(self.selected_datafile.get(), filetype=self.selected_recompiler.get(), sheet_name=self.selected_sheet.get)
=======
        datafile = Datafile(self.selected_datafile.get(), filetype=self.selected_recompiler.get())
>>>>>>> 7966cef (Fixed bug on selecting datafiles)
        test_recompiler = RecompilerMaker.make(datafile)
        self.status_label_timer = 5
        try:
            test_recompiler.read_datafile()
            test_recompiler.reformat()
            test_recompiler.export()
            self.recompile_status['text'] = "Recompile Success!"
        except:
            self.recompile_status['text'] = "Recompiling Failed!"

def main():
    gui = GUI()

if __name__ == "__main__":
    main()