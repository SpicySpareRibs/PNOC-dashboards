import tkinter as tk
from tkinter import ttk, filedialog as fd
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
        self.root.title("Datafile Recompiler")
        self.root.geometry("750x250")
        self.root.resizable(0,0)

        # Initialize widgets
        select_recompiler_label = ttk.Label(self.root, font = ("Times New Roman", 10))
        select_recompiler_label['text'] = "Select the Recompiler Type :"
        select_recompiler_label.grid(column = 0, row = 0, sticky='W') 

        self.selected_recompiler = tk.StringVar()
        select_recompiler_combobox = ttk.Combobox(self.root, width = 27, textvariable = self.selected_recompiler, postcommand = self.reset_status)
        select_recompiler_combobox['values'] = get_args(filetype)
        select_recompiler_combobox.grid(column = 1, row = 0, sticky='W')
        select_recompiler_combobox.current(0)

        select_datafile_label = ttk.Label(self.root, font = ("Times New Roman", 10))
        select_datafile_label['text'] = "Select the datafile: "
        select_datafile_label.grid(column = 0, row = 1, sticky='W') 

        self.datafile_filepath = ""
        select_datafile_button = ttk.Button(self.root, text="Browse", command=self.get_datafile)
        select_datafile_button.grid(column = 1, row = 1, sticky='W')

        self.datafile_filepath_label = ttk.Label(self.root, font = ("Times New Roman", 10), wraplength = 500, justify = 'left')
        self.datafile_filepath_label['text'] = "Datafile Filepath: "
        self.datafile_filepath_label.grid(column = 2, row = 1, sticky='W')

        recompile_button = ttk.Button(self.root, text="Recompile", command=self.recompile)
        recompile_button.grid(column = 1, row = 3, sticky='W')

        self.recompile_status = ttk.Label(self.root, font = ("Times New Roman", 10))
        self.recompile_status['text'] = ""
        self.recompile_status.grid(column = 2, row = 3, sticky='W')

        self.root.mainloop()
    
    def reset_status(self) -> None:
        self.recompile_status['text'] = ""

    def get_datafile(self) -> None:
        self.datafile_filepath = fd.askopenfilename()
        self.datafile_filepath_label['text'] = "Datafile Filepath: " + self.datafile_filepath
        self.reset_status()
    
    def recompile(self) -> None:
        datafile = Datafile(self.datafile_filepath, filetype=self.selected_recompiler.get())
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