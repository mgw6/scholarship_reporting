"""
MacGregor Winegard
11/22/2021

This program groups the scholarships by people who are supposed to recieve a thank you packet.
This makes it so that we can see who is supposed to recieve multiple packets.
The number of individuals in this boat will be small, but there are definitely some. 

Desired output:
Donor name (already in), Spouse name, VP Salutation, addresses, schol_list

"""

import pandas as pd

class XL_work:
    def get_xl_sheet(file_path):
        return pd.read_excel(file_path)
        
if __name__ == '__main__':
    
    recip_list = XL_work.get_xl_sheet("XL_Sheets/merge_output_cleaned_11.22.21.xlsx")
    working = recip_list.copy() #Load df and get it ready to go.
    
    
    prim_name = "Primary Name (Person) (Person)" #The column we are grouping around
    working = working.sort_values(
        by = prim_name,
        ) #Sorts them by the donor recieving the letter
    groups = working.groupby(prim_name)
    #then groups them by donor
    #In retropsect the sort is unnecessary
    
    
    #Set all parallel lists, this is the info we want from the main
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
    
    
    #Loop through each donor
    for person, frame in groups:
        
        """
        #if this donor only gave to one scholarship skip the rest
        Since this was most of the people this saves a lot of processing time
        and cuts down the output
        """
        if len(frame.index) == 1:
            continue 
            
        """
        Matching up columns with the cor lists
        This first block is all of the donors'
        mailing inforation
        """
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
        
        """
        This next block puts all of the scholarships they have contributed 
        to into one column. Because they have donated to different numbers of 
        you can't say its a set number. 
        
        Delimeter for the scholarship names was set as "-" because some
        scholarship names had commas in them. I then split around the delimeter
        in Excel later. 
        """
        temp_schols = ""
        for row in frame.iterrows():
            temp_schols = temp_schols + row[1][0] + "_"
        schol_ls.append(temp_schols)
        
    
    
    
    #Make a pandas df
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
    
    #We will use the seperator in Google sheets to break the scholarships into multiple rows
    
    
    