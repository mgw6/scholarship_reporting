"""
MacGregor Winegard
11/8/2021
This takes the output from group_recips and adds the recipients 
to the merge_output_cleaned file
"""

import pandas as pd
import numpy as np

class XL_work:
    def get_xl_sheet(file_path):
        return pd.read_excel(file_path)


if __name__ == '__main__':
    recip_list = XL_work.get_xl_sheet("XL_Sheets/group_output.xlsx")
    
    primary_sheet = XL_work.get_xl_sheet("XL_Sheets/merge_output_cleaned.xlsx")
    
    name_in_primay = []
    name_in_recip_list = []
    fuzz_score = []
    recips = []
    
    
    
    for row in primary_sheet.iterrows():
        
        #Direct comparison
        if row[1][0] in recip_list['Schol_name'].values:
            name_in_primay.append(row[1][0])
            fuzz_score.append(100)
            recip_row = recip_list.loc[recip_list['Schol_name'] == row[1][0]]
            name_in_recip_list.append(
                recip_row["Schol_name"].values[0]
            )
            recips.append(
                recip_row["Recipients"].values[0]
            )
        
        #Get Fuzzy
        #elif    
        
        else:
            name_in_primay.append(row[1][0])
            fuzz_score.append(0)
            name_in_recip_list.append(pd.NA)
            recips.append(pd.NA)


print(len(recips))
print(len(name_in_primay))
print(len(name_in_recip_list))
print(len(fuzz_score))







