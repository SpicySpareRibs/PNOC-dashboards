import pandas
from bisect import bisect_left
from typing import Protocol,Literal
filetype = Literal["Training Summary Report", "Staffing Summary","Plantilla of Personnel"]


class Datafile:
    def __init__(self, filepath,filetype: filetype):
        self.filepath = filepath
        self.filetype = filetype



class Recompiler(Protocol):
    def read_datafile(self):
        pass

    def get_headers(self):
        pass

    def set_constraints(self):
        pass

    def reformat(self):
        pass
    def export(self):
        pass

class RecompilerMaker():
    def make(datafile: Datafile):
        if datafile.filetype == "Training Summary Report":
            return TrainingSummaryRecompiler(datafile)
        if datafile.filetype == "Plantilla of Personnel":
            return PlantillaOfPersonnelCompiler(datafile)

class TrainingSummaryRecompiler(Recompiler):
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

class PlantillaOfPersonnelCompiler(Recompiler):
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
#test instructions, outputs training_summary_tables.xlsx
# datafile = Datafile("HR\Mock_TRAINING SUMMARY REPORT.xlsx","Training Summary Report")
# test_recompiler = RecompilerMaker.make(datafile=datafile)
# test_recompiler.get_headers()
# test_recompiler.read_datafile()
# test_recompiler.reformat()
# test_recompiler.export()

datafile2 = Datafile("HR\Mock PLANTILLA OF PERSONNEL AND SALARY ADJUSTMENT - BLANK.xlsx","Plantilla of Personnel")
test_recompiler2 = RecompilerMaker.make(datafile=datafile2)
test_recompiler2.read_datafile()
test_recompiler2.reformat()
test_recompiler2.export()