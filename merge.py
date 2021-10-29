#MacGregor Winegard
#10/28/2021 
#This is a quick automation of Scholarship reporting
#Merges FY21 Master with Endowed Scholarship Relationships

import pandas as pd


class XL_work:
    def get_xl_sheet(file_path):
        return pd.read_excel(file_path)
    


if __name__ == '__main__':

    FY21_Primary = XL_work.get_xl_sheet("XL_Sheets/FY21 Scholarship Reporting COPY.xlsx")
    FY21_Primary = FY21_Primary.drop(columns = FY21_Primary.columns[22:]) #remove extra columns
    
    """
    for column in FY21_Primary.columns:
        FY21_Primary.loc[FY21_Primary[column].isna(), column] = "None"
    """
    
    relationships = XL_work.get_xl_sheet("XL_Sheets/Endowed Scholarship Relationships.xlsx")
    
    output = relationships.copy()
    
    #TODO: make new df, then we will append this new one to relationships
    
    
    for row in FY21_Primary.iterrows():
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
        "#TODO",
        full_name,
        full_name,
        pd.NA,
        pd.NA,
        pd.NA,
        pd.NA,
        pd.NA,
        vp_sal, 
        "#TODO",
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
        
        output = output.append(pd.Series(
        new_row,
        index = output.columns),
        ignore_index = True
        )
        


#TODO: Sort data by fund name

output = output.sort_values(
    by = output.columns[0],
    )

output.to_excel("output.xlsx", index = False)
   
    
"""    
print("Schol_name: " + schol_name)
print("Schol_desc: " + schol_desc) 
print("FY20_q4_val: " + str(FY20_q4_val)) 
print("primary_ID: " + str(primary_ID)) 
print("full_name: " + full_name) 
print("spouse: " + spouse) 
print("vp_sal: " + vp_sal) 
print("address_1: " + address_1)
print("address_2: " + address_2) 
print("address_3: " + address_3) 
print("city: " + city) 
print("state: " + state) 
print("zip_code: " + zip_code) 
"""


"""
        relationships.loc[len(relationships)] = [
        schol_name,
        schol_desc,
        "N/A",
        FY20_q4_val,
        "N/A",
        "N/A",
        primary_ID,
        "Yes, NOT this year",
        "None",
        "#TODO",
        full_name,
        full_name,
        "None",
        vp_sal, 
        "#TODO",
        spouse,
        "None",
        address_1,
        address_2,
        address_3,
        city,
        state,
        zip_code,
        ]
        """
        