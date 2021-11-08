#MacGregor Winegard
#11/5/2021 
# This merges clean output with the 

import pandas as pd

# This may be a lifeline:
# https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.from_dict.html

class XL_work:
    def get_xl_sheet(file_path):
        return pd.read_excel(file_path)
    
    
if __name__ == '__main__':
    recip_list = XL_work.get_xl_sheet("XL_Sheets/Advancement Report 10.22.2021.xlsx")
    schol_name = "Description"
    
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
    
    test.to_excel("group_output.xlsx", index = False)
    
    # Using this output, I this in Excel:
    # https://www.adinstruments.com/support/knowledge-base/how-can-comma-separated-list-be-converted-cells-column-lt
    
    #Next, make a third script to do the fuzzy merge of this list into the google sheet
    
    
    
    
    