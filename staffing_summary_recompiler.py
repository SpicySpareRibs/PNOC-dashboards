from typing import Protocol
from recompiler import Datafile, Recompiler
import pandas as pd

class StaffingSummaryRecompiler():
    def __init__(self, datafile:Datafile):
        self._datafile = datafile
        self._dataframe = None
        
    
    
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

    def fill_categories(self,df: pd.DataFrame,category_level):
        category_row = df.columns.get_level_values(category_level).to_list()
        category_index_dict = {}
        for i, name in enumerate(category_row):
            if name not in category_index_dict:
                category_index_dict[name] = [i]
            else: 
                category_index_dict[name].append(i)
        return category_index_dict

    
    def read_datafile(self):
        
        self._dataframe = pd.read_excel(self._datafile.filepath,header=[3,4,5,6],sheet_name=self._datafile.sheet_name) 
        

    def get_headers(self):
        pass
        #list down the relevant headers from the original datafile (can rename)
        #convert them to dictionary keys
    def set_constraints(self):
        pass
        #TBD assign pass for now
    def reformat(self):
        df = self._dataframe
        employee_categories = self.fill_categories(self._dataframe,1)
        all_headers = self.cascade_headers(self._dataframe,[0,1,2],3)
        
        df = df.droplevel([i for i in range(0,df.columns.nlevels - 1)], axis= 1)
        
        df.columns = all_headers
        
        ##next instructions are to reassign the categories and their columns as separate dataframes
        print(employee_categories.keys())
        dataframe_dicts = dict.fromkeys([c for c in employee_categories.keys() if "Unnamed" not in c])
        for c in dataframe_dicts.keys():
            dataframe_dicts[c] = df.iloc[:,[0] + employee_categories[c]]
            dataframe_dicts[c] = dataframe_dicts[c].drop(axis = 1,labels = ["Total"])
            
            dataframe_dicts[c] = dataframe_dicts[c].drop(dataframe_dicts[c].tail(2).index,axis = 0)
            dataframe_dicts[c]["Employment Category"] = c
            #print(dataframe_dicts[c])
            if c == "Contract of Service and/or Job Order":
                
                dataframe_dicts[c].insert(column = 'Authorized',value = "",loc = 1) 
                dataframe_dicts[c].insert(column = "Unfilled",value = "",loc = 4)
                
        frames = [f for f in dataframe_dicts.values()]

        combined_frame = pd.concat(frames)   
        
        self._dataframe = combined_frame
    
    def export(self):
        pass
        with pd.ExcelWriter('staffing_summary.xlsx') as writer:     
            self._dataframe.to_excel(writer,sheet_name="table",index=False)
            self._dataframe.transpose().to_excel(writer,sheet_name="transposed",index=False)
