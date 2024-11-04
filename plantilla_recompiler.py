from typing import Protocol
from recompiler import Datafile,Recompiler
from bisect import bisect_left
import pandas
class PlantillaOfPersonnelCompiler():
    def __init__(self,datafile: Datafile):
        self._datafile = datafile
        self._dataframe = None
        self._employees = None
        self._header_constraints = None
        self._departments = None
    def get_headers(self):
        self._header_constraints = dict.fromkeys([
            "Item No.",
            "Position Title",
            "Job Grade",
            "Salary Step",
            "Previous Salary",
            "Adjusted Salary",
            "Salary Difference",
            "Name of Incumbent",
            "Date of Birth",
            "Tax Identification Number",
            "Date of Original Appointment",
            "Date of Last Promotion",
            "Status of Appointment",
            "REMARKS"
        ])
    def set_constraints(self):
        pass
    def read_datafile(self):
        
        self._dataframe = pandas.read_excel(self._datafile.filepath,header=7)
        df = self._dataframe
        #shorten columns by renaming
        df = df.rename(columns={
            df.columns[1]:"Item No.",
            df.columns[5]:"Previous Salary",
            df.columns[6]:"Adjusted Salary",
            df.columns[7]:"Salary Difference",
            
        })
        #merged cell issue: dept/office label is present in merged cell in col A and B, since col A is to be dropped, we have to copy the 
        #label to col B
        copy_row_index = df.loc[df[df.columns[0]] == "OFFICE OF THE PRESIDENT"].index[0]
        df.loc[copy_row_index,"Item No."] = 'OFFICE OF THE PRESIDENT'
        
        #drop col A
        df = df.drop(columns=[df.columns[0]])
        
        #drop ALL empty rows in between
        df = df.dropna(how='all')
        
        

        #drop row containg (1), (2), (3) etc.
        df = df.drop(df.index[[0]])
        
        #drop rows that contain "Total" in Item Number column
        df = df.drop(df[df["Item No."].str.contains("Total") == True].index)
        df = df.drop(df[(df["Item No."] == "Prepared / Certified Correct") | (df["Item No."] == " Officer-in-Charge, Personnel Services Division   " )].index)
        
        
        self._dataframe = df

    def fill_department_col(self,row_index,dept_dict):
        dept_indexes = list(dept_dict.keys())
        i = bisect_left(dept_indexes,row_index)
        print(i)
        return dept_dict[dept_indexes[i-1]]

    def reformat(self):
        df = self._dataframe

        dept_rows = df[df['Item No.'].apply(lambda x: isinstance(x,str))]
        
        
        
        df = df.drop(dept_rows.index)
        
        df = df.dropna(axis=0,subset=('Item No.'))
        #print(df)
        departments = dept_rows['Item No.']
        print(departments.to_dict())
        
        #add new column for department/Office label
        #Since the original datafile sectioned employee entries by department, we can use the original row indexes to fill in the department/office columns
        #to use the row indexes, we create a temporary column,
        #and use the df[colname].apply function to map the department/office value of a row given its index as follows
        df["tmp"] = df.index
        df["Department/Office"] = df["tmp"].apply(self.fill_department_col, args=(departments.to_dict(),))
        departments = departments.rename("Department Name")
        df = df.drop(columns = ["tmp"])
        print(df)


        self._dataframe = df
        self._departments = departments
    def export(self):
        with pandas.ExcelWriter('regular_employee_table.xlsx') as writer:
            self._dataframe.to_excel(writer,sheet_name="employees",index=False)
            self._departments.to_excel(writer,sheet_name="departments",index=False)