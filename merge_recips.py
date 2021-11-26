"""
MacGregor Winegard
11/8/2021
This takes the output from group_recips and adds the recipients 
to the merge_output_cleaned file

This was my first time using theFuzz package. Its a fun easy package to use. 
Documentation isn't as great as pandas or numpy
but its a realtively simple package so its easy enought to figure out. 
"""

import pandas as pd
import numpy as np
from thefuzz import process



class XL_work: #This was copied from some pd work I did a long time ago
    def get_xl_sheet(file_path):
        return pd.read_excel(file_path)
    


if __name__ == '__main__':
    
    #Extract Sheets
    recip_list = XL_work.get_xl_sheet("XL_Sheets/group_output.xlsx")
    primary_sheet = XL_work.get_xl_sheet("XL_Sheets/merge_output_cleaned.xlsx")
    
    
    """
    These are parallel lists that will turn into each column in the output dfs
    """
    name_in_primay = []
    name_in_recip_list = []
    fuzz_score = []
    recips = []
    
    
    #Loop through each row in our main list of people who recieve letters
    for row in primary_sheet.iterrows():
        
        """
        If there is a direct match, then we plop it right in and don't even call theFuzz.
        """
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
        
        
        """
        If there is not a direct match we call theFuzz. 
        - row[1][0] is the name of the scholarship we are currently working on 
        the big primary list of scholarships
        
        - recip_list['Schol_name'].values is all of the scholarships
        we recipients for in the group recipients
        
        The function goes through the list and finds the best match for the given scholarship name.
        It always returns the best match, even if the best match is really bad (meaning its the best of 
        a bunch of bad options
        
        I don't think I will have to use this again, but if I were to I would actually set a threshold
        and say that if the best match is below the threshold then just don't put in any match. 
        I would set this threshold around 90.
        
        When I was going through later, any scholarship that didn't actually have any recipients were always given 
        terrible matches, which made sense because the correct match didn't exist
        """
        else:    
            best_match = process.extractOne(row[1][0], recip_list['Schol_name'].values) 
            
            name_in_primay.append(row[1][0])
            fuzz_score.append(best_match[1])
            name_in_recip_list.append(best_match[0])
            
            recip_row = recip_list.loc[recip_list['Schol_name'] == best_match[0]]
            recips.append(
                recip_row["Recipients"].values[0]
            )
            
            

    
    """
    Convert 4 lists into one np.array
    This became a rot90 + rot180 when it really
    could have been more of a rot -90
    
    Not sure what I was thinking. 
    """
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
