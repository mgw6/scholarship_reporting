"""
MacGregor Winegard
11/22/2021

This program groups the scholarships by people who are supposed to recieve a thank you packet.
This makes it so that we can see who is supposed to recieve multiple packets.
The number of individuals in this boat will be small, but there are definitely some. 
"""

import pandas as pd

class XL_work:
    def get_xl_sheet(file_path):
        return pd.read_excel(file_path)
        
if __name__ == '__main__':
    
    recip_list = XL_work.get_xl_sheet("XL_Sheets/merge_output_cleaned_11.22.21.xlsx")
    working = recip_list.copy() #Load df and get it ready to go.
    
    
    prim_name = "Primary Name (Person) (Person)"
    working = working.sort_values(
        by = prim_name,
        )  #Sorts them into the order we want
    groups = working.groupby(prim_name)
    #then groups them
    
    
    
    donor_name_ls = []
    spouse_ls = []
    vp_sal_ls = []
    street1_ls = []
    street2_ls = []
    street3_ls = []
    city_ls = []
    state_ls = []
    zip_ls = []
    schol_ls = []
    
    for person, frame in groups:
        
        if len(frame.index) == 1:
            continue #We only want the ones with multiple recipients
        
        
        donor_name_ls.append(person)
        #print(frame.iloc[0][0])
        spouse_ls.append(frame.iloc[0][19])
        vp_sal_ls.append(frame.iloc[0][17])
        street1_ls.append(frame.iloc[0][21])
        street2_ls.append(frame.iloc[0][22])
        street3_ls.append(frame.iloc[0][23])
        city_ls.append(frame.iloc[0][24])
        state_ls.append(frame.iloc[0][25])
        zip_ls.append(frame.iloc[0][26])
        
        temp_schols = ""
        
        for row in frame.iterrows():
            temp_schols = temp_schols + row[1][0] + "_"
        
        schol_ls.append(temp_schols)
        
    
    """
    print(
    donor_name_ls,
    spouse_ls,
    vp_sal_ls,
    street1_ls,
    street2_ls,
    street3_ls,
    city_ls,
    state_ls,
    zip_ls,
    sep = '\n\n'
    )    
    exit()
    """
    
    output = pd.DataFrame({
    'Donor Name' : donor_name_ls,
    'Spouse' : spouse_ls,
    'VP Salutation': vp_sal_ls,
    'Street 1' : street1_ls,
    'Street 2' : street2_ls,
    'Street 3' : street3_ls,
    'City' : city_ls,
    'State' : state_ls,
    'Zip' : zip_ls,
    'Scholarships' : schol_ls
    })
    
    output.to_excel("dup_donors.xlsx", index = False)
    print("Output saved")
    