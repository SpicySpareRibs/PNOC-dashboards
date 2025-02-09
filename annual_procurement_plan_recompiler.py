from typing import Protocol
from recompiler import Datafile, Recompiler
import pandas as pd

class PAPSupertaskRecompiler():
    def __init__(self, datafile):
        self._datafile = datafile
        self._df_subtasks = None
        self._df_supertasks = None

    
    def read_datafile(self):
        df_subtasks = pd.read_excel(self._datafile.filepath,header=[3])
        df_subtasks = df_subtasks.dropna(axis = 0, subset=('PMO/End-User'))
        df_subtasks = df_subtasks.dropna(axis = 0, subset=('Schedule of Each Procurement Activity'))
        df_subtasks = df_subtasks.dropna(axis = 1, how='all')
        df_subtasks = df_subtasks.drop(columns=['Unnamed: 11', 'Unnamed: 12'])
        
        self._df_subtasks = df_subtasks
        
        self._df_supertasks = pd.read_excel(self._datafile.filepath,header=[3,4])

    def get_headers(self):
        pass
        #list down the relevant headers from the original datafile (can rename)
        #convert them to dictionary keys
    
    def set_constraints(self):
        pass
        #TBD assign pass for now
    
    def split_code(self, df):
        # Sets the supercode and subcode of the procurement
        subcodes = []
        supercodes = []
        codes = df["Subcode"]
        for code in codes:
            is_subcode_numeric = code[-1].isnumeric()
            split = -1
            if is_subcode_numeric is True:
                while code[split-1].isnumeric() is True:
                    split-=1
            else:
                while code[split-1].isnumeric() is False:
                    split-=1
            
            subcodes.append(code[split:])
            supercodes.append(code[:split])

        df["Subcode"] = subcodes
        df.insert(0, "Supercode", supercodes)

        return (df)

    def split_start_finish(self, df):
        # Splits schedule into starting quarter and end quarter
        quarters = ['1st', '2nd', '3rd', '4th']

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

        return (df)

    def cascade_headers(self,df: pd.DataFrame,top_levels,bottom_level):
        
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

    def reformat(self):
        # Reformat subtasks sheet
        df_subtasks = self._df_subtasks

        df_subtasks = df_subtasks.rename(columns={
            df_subtasks.columns[0]:"Subcode",
            df_subtasks.columns[1]:"Name",
            df_subtasks.columns[2]:"End-user",
            df_subtasks.columns[3]:"Is Early Procurement",
            df_subtasks.columns[5]:"Quarter Start",
            df_subtasks.columns[6]:"Budget Estimate",
        })

        df_subtasks = self.split_start_finish(df_subtasks)
        df_subtasks = self.split_code(df_subtasks)

        self._df_subtasks = df_subtasks

        # Reformat supertasks sheet

        df_supertasks = self._df_supertasks
        all_headers = self.cascade_headers(df_supertasks,[0],1)

        df_supertasks = df_supertasks.droplevel([i for i in range(0,df_supertasks.columns.nlevels - 1)], axis= 1)
        df_supertasks.columns = all_headers

        df_supertasks.rename({'Advertisement / Posting of IB / REI': 'Schedule of Each Procurement Activity'}, axis=1, inplace=True)
        df_supertasks.drop(['Submission / Opening of Bids', 'Notice of Award', 'Contract Signing'], axis=1)


        columns_check = [
            "PMO/End-User",
            "Is this an Early Procurement Activity? (Yes/No)",
            "Mode of Procurement",
            "Schedule of Each Procurement Activity"
        ]
        mask = df_supertasks[columns_check].isna().any(axis=1)
        # Print the 'PMO/End User' column for rows where all `columns_check` are NaN
        
        df_supertasks_cleaned = df_supertasks[mask]
        #print(df_supertasks_cleaned)

        new_df_supertasks = df_supertasks_cleaned.dropna(subset=["Code \n(PAP)"])

        df_supertask = pd.DataFrame()

        df_supertask['supercode'] = new_df_supertasks['Code \n(PAP)']
        df_supertask['name'] = new_df_supertasks['PROCUREMENT PROJECT']
           

        # Apply the function to create 'type' column
        df_supertask['type'] = new_df_supertasks['MOOE'].notna().apply(lambda x: 'MOOE' if x else 'CAPITAL OUTLAY')


        # Print the new DataFrame
        #print(df_supertask)

        # Update the DataFrame with cleaned data
        self._df_supertasks = df_supertask
    
    def export(self):
        with pd.ExcelWriter('annual_procurement_plan.xlsx') as writer:     
            self._df_supertasks.to_excel(writer,sheet_name="supertasks",index=False)
            self._df_subtasks.to_excel(writer,sheet_name="subtasks",index=False)