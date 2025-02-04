from typing import Protocol
from recompiler import Datafile,Recompiler
import pandas
from pandas import Timestamp
import openpyxl
from openpyxl.utils import get_column_letter
import numpy as np
class ProcTaskCompiler():
    def __init__(self,datafile):
        self._datafile = datafile
        self._dataframe = None
        self._header_constraints = None
        self._stages = None
        self._record_stages = None
        self._record = None
        self._record_codes = None

    def get_headers(self):
        pass
    def set_constraints(self):
        pass
    
    def cascade_headers(self,df: pandas.DataFrame,top_levels,bottom_level):
        
        bottom_row = df.columns.get_level_values(bottom_level).to_list()
        
        
        items_per_level = dict((l,df.columns.get_level_values(l).to_list()) for l in top_levels)
        for j in range(0,len(bottom_row)):
                if "Unnamed" in bottom_row[j]:
                    for i in reversed(top_levels):
                        if "Unnamed" not in items_per_level[i][j]:
                            bottom_row[j] = items_per_level[i][j]
                            break
        final_headers = bottom_row
        return final_headers
    def read_datafile(self):

        self._dataframe = pandas.read_excel(self._datafile.filepath,header=[3,4])
        #df = self._dataframe
        #df.columns.get_level_values(1).to_list()
        #print(df.columns.get_level_values(1).to_list())
        
        
    def reformat(self):
        df = self._dataframe
        all_headers = self.cascade_headers(self._dataframe,[0],1)

        df = df.droplevel([i for i in range(0,df.columns.nlevels - 1)], axis= 1)
        df.columns = all_headers

        df.rename({'Advertisement / Posting of IB / REI': 'Schedule of Each Procurement Activity'}, axis=1, inplace=True)
        df.drop(['Submission / Opening of Bids', 'Notice of Award', 'Contract Signing'], axis=1)


        columns_check = [
            "PMO/End-User",
            "Is this an Early Procurement Activity? (Yes/No)",
            "Mode of Procurement",
            "Schedule of Each Procurement Activity"
        ]
        mask = df[columns_check].isna().any(axis=1)
        # Print the 'PMO/End User' column for rows where all `columns_check` are NaN
        
        df_cleaned = df[mask]
        #print(df_cleaned)

        new_df = df_cleaned.dropna(subset=["Code \n(PAP)"])

        df_supertask = pandas.DataFrame()

        df_supertask['supercode'] = new_df['Code \n(PAP)']
        df_supertask['name'] = new_df['PROCUREMENT PROJECT']

       
       
           

        # Apply the function to create 'type' column
        df_supertask['type'] = new_df['MOOE'].notna().apply(lambda x: 'MOOE' if x else 'CAPITAL OUTLAY')


        # Print the new DataFrame
        #print(df_supertask)
    


        # Update the DataFrame with cleaned data
        self._dataframe = df_supertask

        




    def export(self):
        with pandas.ExcelWriter('PAP_tasks.xlsx') as writer:     
            self._dataframe.to_excel(writer,sheet_name="supertasks",index=False)

# Under Dev            
df = Datafile(filepath="Procurement/PNOC 2024 1st Sem Revised APP.xlsx",filetype="Procurement Monitoring Report")
test = ProcTaskCompiler(datafile=df)
test.read_datafile()
test.reformat()
test.export()