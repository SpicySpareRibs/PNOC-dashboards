from recompiler import Datafile, filetype
from plantilla_recompiler import PlantillaOfPersonnelCompiler
from training_summary_recompiler import TrainingSummaryRecompiler
from contract_recompiler import ContractCompiler
from proc_monitorning_report import ProcMonitoringCompiler
from inventory_of_clients_recompiler import InventoryOfClientsRecompiler

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
            case _:
                print("Invalid Input")
                return None

datafile = Datafile("HR\Mock_TRAINING SUMMARY REPORT.xlsx","Training Summary Report")
test_recompiler = RecompilerMaker.make(datafile=datafile)
test_recompiler.get_headers()
test_recompiler.read_datafile()
test_recompiler.reformat()
test_recompiler.export()

datafile2 = Datafile("HR\Mock PLANTILLA OF PERSONNEL AND SALARY ADJUSTMENT - BLANK.xlsx","Plantilla of Personnel")
test_recompiler2 = RecompilerMaker.make(datafile=datafile2)
test_recompiler2.read_datafile()
test_recompiler2.reformat()
test_recompiler2.export()

datafile3 = Datafile("HR\Mock_2024 Contract of Service Employees.xlsx","Contract of Service")
test_recompiler3 = RecompilerMaker.make(datafile=datafile3)
test_recompiler3.read_datafile()
test_recompiler3.export()

datafile4 = Datafile("Project Management\Mock_Inventory_of_Clients - RGB (1).xlsx","Inventory of Clients")
test_recompiler4 = RecompilerMaker.make(datafile=datafile4)
test_recompiler4.read_datafile()
test_recompiler4.reformat()
test_recompiler4.export()