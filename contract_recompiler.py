from typing import Protocol
from recompiler import Datafile,Recompiler
from bisect import bisect_left
import pandas

class ContractCompiler():
    def __init__(self,datafile: Datafile):
        self._datafile = datafile
        self._dataframe = None
        self._position = None
        self._header_constraints = None
    
    def get_headers(self):
        self._header_constraints = dict.fromkeys([
            "NAME OF INCUMBENT",
            "POSITION TITLE",
            "DEPARTMENT/OFFICE",
            "DATE OF BIRTH",
            "CONTACT NUMBER",
            "EMAIL ADDRESS",
            "SALARY",
            "MONTHLY RATE",
            "START DATE",
            "END DATE",
            "STATUS OF PRE-EMPLOYMENT REQUIREMENTS",
            "STATUS OF CONTRACT",
            "REMARKS"
        ])
    def set_constraints(self):
        pass
    def read_datafile(self):
        
        self._dataframe = pandas.read_excel(self._datafile.filepath,header=2,sheet_name=self._datafile.sheet_name)
        df = self._dataframe
        #renamed some columns based on specified in system requirements doc
        df = df.rename(columns={
            df.columns[2]:"DEPARTMENT NAME",
            df.columns[10]: "PRE-EMPLOYMENT STATUS"   
        })
        
        
        self._dataframe = df


    def reformat(self):
        df = self._dataframe

        #other_df = pandas.read_excel("regular_employee_table.xlsx",header=0)

        #df['POSITION TITLE'] = df['POSITION TITLE'].map(other_df.set_index('POSITION TITLE')['Position Title'])

        #df['DEPARTMENT NAME'] = df['DEPARTMENT NAME'].map(other_df.set_index('DEPARTMENT NAME')['Department/Office'])

        #position = df['POSITION TITLE']
        #departments = df['DEPARTMENT NAME']

        #self.position = position

        #self.departments = departments

        #self._dataframe = df

    def export(self):
        with pandas.ExcelWriter('contract_employee_table.xlsx') as writer:
            self._dataframe.to_excel(writer,sheet_name="contracts",index=False)

        