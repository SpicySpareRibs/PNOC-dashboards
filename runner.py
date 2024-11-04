from recompiler import Datafile, filetype
from plantilla_recompiler import PlantillaOfPersonnelCompiler
from training_summary_recompiler import TrainingSummaryRecompiler

class RecompilerMaker():
    def make(datafile: Datafile):
        if datafile.filetype == "Training Summary Report":
            return TrainingSummaryRecompiler(datafile)
        if datafile.filetype == "Plantilla of Personnel":
            return PlantillaOfPersonnelCompiler(datafile)

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