"""
MacGregor Winegard
11/8/2021
This takes the output from group_recips and adds the recipients 
to the merge_output_cleaned file
"""

import pandas as pd
import numpy as np
from thefuzz import process

class XL_work:
    def get_xl_sheet(file_path):
        return pd.read_excel(file_path)


if __name__ == '__main__':
    
    #Extract Sheets
    recip_list = XL_work.get_xl_sheet("XL_Sheets/group_output.xlsx")
    primary_sheet = XL_work.get_xl_sheet("XL_Sheets/merge_output_cleaned.xlsx")
    
    #Setup lists
    name_in_primay = []
    name_in_recip_list = []
    fuzz_score = []
    recips = []
    
    
    #Loop through each row
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
        
        
        else:    
            best_match = process.extractOne(row[1][0], recip_list['Schol_name'].values) 
            
            name_in_primay.append(row[1][0])
            fuzz_score.append(best_match[1])
            name_in_recip_list.append(best_match[0])
            
            recip_row = recip_list.loc[recip_list['Schol_name'] == best_match[0]]
            recips.append(
                recip_row["Recipients"].values[0]
            )
            
            

    
    #Covnert 4 lists into one np.array
    data_list = np.flipud(np.rot90(
        [
            name_in_primay,
            name_in_recip_list,
            fuzz_score,
            recips,
        ]
 
    ))
    
    
    #Put the array into a pd.df and export to excep1
    output_df = pd.DataFrame(
        data = data_list, 
        columns = [ 
            "name_in_primay",
            "name_in_recip_list",
            "fuzz_score",
            "recips", 
            ]
        )
    output_df.to_excel("merge_recips_output.xlsx", index = False)






