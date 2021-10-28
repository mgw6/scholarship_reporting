#MacGregor Winegard
#10/28/2021 
#This is a quick automation of Scholarship reporting
#Merges FY21 Master with Endowed Scholarship Relationships

import pandas as pd


class XL_work:
    def get_xl_sheet(file_path):
        return pd.read_excel(file_path)
    


if __name__ == '__main__':

    FY21_Primary = XL_work.get_xl_sheet("XL_Sheets/FY21 Scholarship Reporting MASTER LIST 2.4.21 EMR.xlsx")
    #TODO: Change this to my copy
    
    FY21_Primary = FY21_Primary.drop(columns = FY21_Primary.columns[22:]) #remove extra columns
    #TODO: remove those company rows
    #may just brute force
    
    relationships = XL_work.get_xl_sheet("XL_Sheets/Endowed Scholarship Relationships.xlsx")
    
    
    
    for row in FY21_Primary.iterrows():
        print(row)
        exit()
    
    # for row in my copy
    #   Line up the info from my copy into the new relationhsips ones
    #maybe add a column that is something like "only in the old one"
    
    #write to excel sheet