"""
MacGregor Winegard
10/28/2021 
This is a quick automation of Scholarship reporting
Merges FY21 Master with Endowed Scholarship Relationships
"""

import pandas as pd


class XL_work: #This was copied from some pd work I did a long time ago
    def get_xl_sheet(file_path):
        return pd.read_excel(file_path)
    


if __name__ == '__main__':

    FY21_Primary = XL_work.get_xl_sheet("XL_Sheets/FY21 Scholarship Reporting COPY.xlsx")
    FY21_Primary = FY21_Primary.drop(columns = FY21_Primary.columns[22:]) #remove extra columns
    
    relationships = XL_work.get_xl_sheet("XL_Sheets/Endowed Scholarship Relationships.xlsx")
    
    output = relationships.copy()
    
    
    for row in FY21_Primary.iterrows(): 
        """
        Loop through the rows in last years list
        This first part just pulls out info from the specified columns
        """
        schol_name = row[1][0]
        schol_desc = row[1][1]  
        FY20_q4_val = row[1][2]
        
        try:
            primary_ID = int(row[1][5])
        except:
            primary_ID = pd.NA
        
        
        full_name = row[1][6]
        spouse = row[1][7]
        vp_sal = row[1][8]
        address_1 = row[1][9]
        address_2 = row[1][10]
        address_3 = row[1][11]
        city = row[1][12]
        state = row[1][13]
        zip_code = row[1][14]
        
        try:
            name_list = str(full_name).split(' ')
            sort_name = name_list[-1].upper() + ", " + (' '.join(name_list[1:-1])).upper()
            last_name = ' '.join(name_list[2:])
        except:
            sort_name = pd.NA
            last_name = pd.NA
        
        
        
        """
        Now we make a new row as a list
        """
        new_row = [
        schol_name,
        schol_desc,
        pd.NA,
        FY20_q4_val,
        pd.NA,
        pd.NA,
        primary_ID,
        "ONLY",
        pd.NA,
        sort_name,
        full_name,
        full_name,
        pd.NA,
        pd.NA,
        pd.NA,
        pd.NA,
        pd.NA,
        vp_sal, 
        last_name,
        spouse,
        pd.NA,
        address_1,
        address_2,
        address_3,
        city,
        state,
        zip_code,
        pd.NA,
        pd.NA,
        pd.NA,
        pd.NA,
        pd.NA,
        pd.NA,
        pd.NA,
        ]
        
        
        """
        Append new row to output df
        I know I could have done that in way less lines, 
        but in making sure that I was lining up the columns correctly from one sheet to the other
        I just pulled them out as variables and then printed the variables.
        So its not the most memory or time efficient way but it does the job.
        I also was not terribly worried about memory or time because this was a relatively small data set. 
        Correctness and certainty of it was way more important. 
        """
        
        output = output.append(pd.Series(
        new_row,
        index = output.columns),
        ignore_index = True
        )

    #Sort by Scholarship name
    output = output.sort_values(
        by = output.columns[0],
        )
        
        
    output.to_excel("merge_output.xlsx", index = False)
    """
    Output is the relationships sheet, with the 
    values that were only on FY 21 Master added. 
    """