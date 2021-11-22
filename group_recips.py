"""
MacGregor Winegard
11/5/2021 
This program goes through the list of Scholarship recipients and groups them by scholarship. 
Output is such that column one is the scholarship and column two is each recipient, delimited by '.'

I copy and pasted this into the main excel file and then used the split text option to break them into multiple columns
"""

import pandas as pd

# This may be a lifeline:
# https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.from_dict.html

class XL_work:
    def get_xl_sheet(file_path):
        return pd.read_excel(file_path)
    
    
if __name__ == '__main__':
    recip_list = XL_work.get_xl_sheet("XL_Sheets/Update 11.17.2021.xlsx")
    #TODO: Make a Tkinter option for this. 
    
    
    schol_name = "RFRBASE_FUND_TITLE_LONG" #This was changed on the 11.17 update. 
    
    name_list = ['First Name', 'Last Name']
    
    working = recip_list[[name_list[0], name_list[1], schol_name]].copy()
    
    
    working = working.sort_values(
    by = schol_name,
    )
    
    groups = working.groupby(schol_name)
    
    new_dict = {}
    
    for schol, frame in groups:
        #print("Name: " + str(schol))
        
        temp_list = []
        
        for row in frame.iterrows():
            temp_list.append(row[1][0]  + " " + row[1][1])
            #print(row[1][0], row[1][1] )
        
        new_dict[schol] = temp_list
    
    list_for_df = []
    
    for key, value in new_dict.items():
        list_for_df.append([key, ', '.join(value)])
        #print(key, " : ", value)
        #print()
    
    
    
    test = pd.DataFrame.from_dict(list_for_df)
    
    test.to_excel("11.17_update_group.xlsx", index = False)
    print("Output saved")
    
    # Using this output, I this in Excel:
    # https://www.adinstruments.com/support/knowledge-base/how-can-comma-separated-list-be-converted-cells-column-lt
    
    #Next, make a third script to do the fuzzy merge of this list into the google sheet
    
    
    
    
    