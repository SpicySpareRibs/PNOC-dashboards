from typing import Protocol
from recompiler import Datafile,Recompiler
import pandas
from pandas import Timestamp
import openpyxl
from openpyxl.utils import get_column_letter
import numpy as np
class ProcMonitoringCompiler():
    def __init__(self,datafile):
        self._datafile = datafile
        self._dataframe = None
        self._header_constraints = None
        self._stages = None
        self._record_stages = None
        self._record = None
        self._record_code = None

    def get_headers(self):
        pass
    def set_constraints(self):
        pass
    def get_hidden_cols(self,max):
        workbook = openpyxl.load_workbook(self._datafile.filepath)
        ws = workbook["RMB-PMR 2024 "]
        
        max_col = ws.max_column
        cols = [get_column_letter(i) for i in range(1, max_col+1)]

        hidden_cols = []
        last_hidden = 0
        for i, col in enumerate(cols):
            if i >= max:
                break
        # Column is hidden
            if ws.column_dimensions[col].hidden:
                hidden_cols.append(i)
                # Last column in the hidden group
                last_hidden = ws.column_dimensions[col].max
        # Appending column if more columns in the group
            elif i+1 <= last_hidden:
                hidden_cols.append(i)
            
        
        return hidden_cols 
    def replace_header(self,level0,level1):
            pass
    def replace_date(self, timestamp):
        
        try: 
            
            
            return timestamp.date() 
        except:
            print("Error: input is not a timestamp")
            return timestamp
    def read_datafile(self):
        df = pandas.read_excel(self._datafile.filepath,sheet_name="RMB-PMR 2024 ",header=[6,7],)
        hidden_cols = self.get_hidden_cols(len(df.columns))
        
        level0 = df.columns.get_level_values(0)
        
        new_cols = df.columns.get_level_values(1).to_list()
        for i in range(len(new_cols)):
            if "Unnamed" in new_cols[i]:
                new_cols[i] = level0[i]
        
        df.columns = df.columns.droplevel(0)
        df.columns = new_cols
        
        df.drop(df.columns[hidden_cols],axis=1,inplace=True)

        for i, col in enumerate(df.columns):
            print(f"{i}: {col}")
    
        df[df.columns[5]] = df.apply(lambda x: self.replace_date(x[df.columns[5]]),axis = 1)

        self._datafile = df

    

    def reformat(self):
        pass
    def export(self):
        with pandas.ExcelWriter('PMR.xlsx',datetime_format="MM-DD-YYYY",date_format="YYYY-MM-DD") as writer:     
            self._datafile.to_excel(writer,sheet_name="test",index=False)

# code still in progress          
# df = Datafile(filepath="Procurement\PMR.xlsx",filetype="Procurement Monitoring Report")
# test = ProcMonitoringCompiler(datafile=df)
# test.read_datafile()
# test.export()
