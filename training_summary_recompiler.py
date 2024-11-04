from typing import Protocol
from recompiler import Datafile,Recompiler
import pandas
class TrainingSummaryRecompiler():
    def __init__(self,datafile: Datafile):
        self._datafile = datafile
        self._dataframe = None
        self._header_constraints = None 
        self._trainees = None
        self._sessions = None
    def get_headers(self):
        #for each datafile, we will have to copy/list the column headers and convert them as dict keys
        self._header_constraints = dict.fromkeys([
        "No",
        "Employee Name",
        "Department/Office",
        "CODE",
        "Program Title",
        "Sponsor",
        "Duration",
        "Quarter",
        "No. of days",
        "# HRS",
        " Sem. Fee "])
    def set_constraints(self):
        #once the dictionary keys are obtained, we apply a listing of constraints (similar to the contraints in SQL)
        #this is to be used later when we browse each column for data validation in reformat()
        pass

    def read_datafile(self):
        #first arg is filepath, header argument is to specify which row column headers start
        self._dataframe = pandas.read_excel(self._datafile.filepath,header=2)
        
    def reformat(self):
        #listing is done to limit the columns included in each normalized table
        self._trainees = self._dataframe[["Employee Name","Department/Office","CODE","Program Title","Sponsor"]]
        self._sessions = self._dataframe[["Program Title","Duration","Quarter","No. of days","# HRS","Sem. Fee"]]

        #to do: ask if training session details are unique per trainee, if not, we can apply group bys/distincts
        #we also apply here the rules for constraints by browsing row by row. If a specific row does not follow 
        #constraints, we print an error code and NOT include in into the final dataframe
        print(self._sessions)
    def export(self):
        #to create the new excel file, call the ff:
        #index = False means no additional index column would be included in the new excel file
        with pandas.ExcelWriter('training_summary_table.xlsx') as writer:     
            self._trainees.to_excel(writer,sheet_name="trainees",index=False)
            self._sessions.to_excel(writer,sheet_name="sessions",index=False)


