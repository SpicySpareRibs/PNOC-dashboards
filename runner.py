from recompiler import Datafile, filetype
from plantilla_recompiler import PlantillaOfPersonnelCompiler
from training_summary_recompiler import TrainingSummaryRecompiler
from contract_recompiler import ContractCompiler

class RecompilerMaker():
    def make(datafile: Datafile):
        if datafile.filetype == "Training Summary Report":
            return TrainingSummaryRecompiler(datafile)
        if datafile.filetype == "Plantilla of Personnel":
            return PlantillaOfPersonnelCompiler(datafile)
        if datafile.filetype == "Contract of Service":
            return ContractCompiler(datafile)

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

datafile3 = Datafile("HR\Mock_2024 Contract of Service.xlsx","Contract of Service")
test_recompiler3 = RecompilerMaker.make(datafile=datafile3)
test_recompiler3.read_datafile()
test_recompiler3.export()