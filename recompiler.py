import pandas
from bisect import bisect_left
from typing import Protocol,Literal



filetype = Literal["Training Summary Report", 
                   "Staffing Summary",
                   "Plantilla of Personnel", 
                   "Contract of Service",
                   "Procurement Monitoring Report"]



class Datafile:
    def __init__(self, filepath,filetype: filetype):
        self.filepath = filepath
        self.filetype = filetype



class Recompiler(Protocol):
    def read_datafile(self):
        pass
        #use pandas.read_excel() to convert the original datafile to a pandas dataframe
        #optional here is to do preliminary reformatting
    def get_headers(self):
        pass
        #list down the relevant headers from the original datafile (can rename)
        #convert them to dictionary keys
    def set_constraints(self):
        pass
        #TBD assign pass for now
    def reformat(self):
        pass
        #change dataframe to output tables as per schema
        #methods can vary for difficult datafiles
    def export(self):
        pass
        #use pandas.ExcelWriter() and .to_excel() to output a new excel file
        #1 excel file for now is to one recompiler, 1 excel file can contain multiple sheets as tables


