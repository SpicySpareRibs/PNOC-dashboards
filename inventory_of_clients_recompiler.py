from typing import Protocol
from recompiler import Datafile,Recompiler
from bisect import bisect_left
import pandas
import numpy as np

class InventoryOfClientsRecompiler:
    def __init__(self, datafile:Datafile):
        self._datafile = datafile
        self._dataframe = None
        self._clients = None
        self._current = None
        self._header_constraints = None
    
    def read_datafile(self):
        df = pandas.read_excel(self._datafile.filepath,header=0,sheet_name=self._datafile.sheet_name)
        df.columns = pandas.Index([
            'Government Agency',
            'Status',
            'kWp',
            'Total ABC (Php)',
            'CAPEX per kW (Php)',
            'Contact Person',
            'Activity',
            'Last Contact',
            'Focal Person',
            'Remarks',
            'Files',
            'RES',
            'Notes',
        ], dtype = 'object')
        
        df = df.dropna(axis = 0, subset=('Government Agency'))
        df = df.replace('#DIV/0!',np.nan)

        self._dataframe = df

        #use pandas.read_excel() to convert the original datafile to a pandas dataframe
        #optional here is to do preliminary reformatting


    def get_headers(self):
        pass

    def set_constraints(self):
        pass

    def reformat(self):
        clients = self._dataframe
        currentprojects = clients.drop(columns = ['Contact Person','Activity','Last Contact','Focal Person','Remarks','Files','RES','Notes',])
        clients = clients.drop(columns = ["kWp", "Total ABC (Php)", "CAPEX per kW (Php)"])
        
        currentprojects = currentprojects[currentprojects['Status'] != 'Touch Base']
        currentprojects = currentprojects[currentprojects['Status'] != 'Discontinued']
        
        self._clients = clients
        self._currentprojects = currentprojects

    def export(self):
        with pandas.ExcelWriter('inventory_of_clients.xlsx', datetime_format="MM-DD-YYYY", date_format="MM-DD-YYYY") as writer:
            self._clients.to_excel(writer,sheet_name="clients",index=False)
            self._currentprojects.to_excel(writer,sheet_name="current projects",index=False)
            self._dataframe.to_excel(writer,sheet_name="inventory",index=False)
        #use pandas.ExcelWriter() and .to_excel() to output a new excel file
        #1 excel file for now is to one recompiler, 1 excel file can contain multiple sheets as tables