from typing import Protocol
from recompiler import Datafile, Recompiler
import pandas as pd

class PAPSupertaskRecompiler():
    def __init__(self, datafile):
        self._datafile = datafile
        self._dataframe = None
    
    def read_datafile(self):
        df = pd.read_excel(self._datafile.filepath,header=[3])
        df = df.dropna(axis = 0, subset=('PMO/End-User'))
        df = df.dropna(axis = 0, subset=('Schedule of Each Procurement Activity'))
        df = df.dropna(axis = 1, how='all')
        df = df.drop(columns=['Unnamed: 11', 'Unnamed: 12'])
        
        self._dataframe = df
        

    def get_headers(self):
        pass
        #list down the relevant headers from the original datafile (can rename)
        #convert them to dictionary keys
    
    def set_constraints(self):
        pass
        #TBD assign pass for now
    
    def split_code(self):
        # Sets the supercode and subcode of the procurement
        df = self._dataframe
        
        code = df["Subcode"]
        df.insert(0, "Supercode", code)

        self._dataframe = df

    def split_start_finish(self):
        # Splits schedule into starting quarter and end quarter
        quarters = ['1st', '2nd', '3rd', '4th']
        df = self._dataframe

        schedule = df["Quarter Start"]
        quarter_start_index = df.columns.get_loc("Quarter Start")
        quarter_start = []
        quarter_finish = []
        for entry in schedule:
            quarter_start.append(entry.strip()[:3])
            if len(entry) > 11:
              substrings = entry.strip()[3:].split(" ")
              for substring in substrings:
                  if len(substring) == 4: # check for year
                      quarter_start[-1] = quarter_start[-1] + ' (' + substring + ')'
                  if substring in quarters: # if end quarter is found
                      if len(substrings[-1]) == 4: # check for year
                          quarter_finish.append(substring + ' (' + substrings[-1] + ')')
                      else:
                          quarter_finish.append(substring)
                      break
                  if substring == substrings[-1]: # if end quarter is not found, set to quarter start
                      quarter_finish.append(entry.strip()[:3])
            else:
              quarter_finish.append(entry[:3])

        df["Quarter Start"] = quarter_start
        df.insert(quarter_start_index+1, "Quarter Finish", quarter_finish)
        
    def reformat(self):
        df = self._dataframe

        df = df.rename(columns={
            df.columns[0]:"Subcode",
            df.columns[1]:"Name",
            df.columns[2]:"End-user",
            df.columns[3]:"Is Early Procurement",
            df.columns[5]:"Quarter Start",
            df.columns[6]:"Budget Estimate",
        })

        self._dataframe = df

        self.split_start_finish()
        self.split_code()
    
    def export(self):
        with pd.ExcelWriter('app_supertasks.xlsx') as writer:     
            self._dataframe.to_excel(writer,sheet_name="table",index=False)