"""
MacGregor Winegard
11/5/2021 
This program goes through the list of Scholarship recipients and groups them by scholarship. 
Output is such that column one is the scholarship and column two is each recipient, delimited by ','

I copy and pasted this into the main excel file and then used the split text option to break them into multiple columns
"""

import pandas as pd
import tkinter.filedialog #tkinter was being a pain, I'm not sure why
import tkinter as tk
import os


# This may be a lifeline:
# https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.from_dict.html

class XL_work:
    def get_xl_sheet(file_path): #Copied from Pandas work I did a year ago
        return pd.read_excel(file_path)
    
    
if __name__ == '__main__':
    
    
    #Gives the user a GUI to get the file
    root = tk.Tk()
    root.withdraw()
    file_path = tkinter.filedialog.askopenfilename(filetypes = [('Excel Files', '*.xlsx')],
                                            initialdir = "C:/Users/User/exec_intern/scholarship_reporting",
                                            title = "Select The excel sheet" 
                                            ) #This opens the file selector  
    
    recip_list = XL_work.get_xl_sheet(file_path)
    
    """
    Name of the scholarship column
    This was changed on the 11.17 update. 
    """
    schol_name = "RFRBASE_FUND_TITLE_LONG" 
    
    name_list = ['First Name', 'Last Name'] 
    #This is the names of the columns we want to pull from
    
    
    working = recip_list[[name_list[0], name_list[1], schol_name]].copy()
    #To save memory and time I copied only the 3 columns I needed info from
    
    working = working.sort_values(
    by = schol_name,
    )
    """
    Sort by scholarship name
    I think it came this way anyways, but just to be sure I threw this in there.
    In retrospect though it probably didn't matter since I grouped it anyways
    """
    
    groups = working.groupby(schol_name) 
    #Grouped by scholarship name
    
    new_dict = {}
    """
    We're pulling this into a dictionary where the key is the scholarship name
    And the value is a string that has all of the recpients
    
    """
    
    
    for schol, frame in groups:
        #Go through all of the groups
        #print("Name: " + str(schol))
        
        
        """
        Every value in this list is the first and last names of every person 
        who recieved this scholarship.
        i.e.: ["MacGregor Winegard", "Derek Jeter", "John Mayer"]
        """
        temp_list = []
        for row in frame.iterrows():
            temp_list.append(row[1][0]  + " " + row[1][1])
            #print(row[1][0], row[1][1] )
        
        #Now add this key:value pair to the dict
        new_dict[schol] = temp_list
    
    
    """
    Make a list that will be dropped right into a pandas df. 
    Format of the list is 
    
    [
    ["The Fisher Scholarship", "MacGregor Winegard,"]
    ["The Service Scholarship", "MacGregor Winegard, Clayton Kershaw, Walker Beuhler"]
    etc 
    ]
    """
    list_for_df = []
    for key, value in new_dict.items():
        list_for_df.append([key, ', '.join(value)])
        #print(key, " : ", value)
        #print()
    
    
    #Make a pd.DataFrame from this list of lists
    test = pd.DataFrame.from_dict(list_for_df)
    
    
    output_name = os.path.splitext(os.path.basename(file_path))[0] + "_grouped.xlsx"
    
    test.to_excel(output_name, index = False)
    print("Output saved")
    
    # Using this output, I this in Excel:
    # https://www.adinstruments.com/support/knowledge-base/how-can-comma-separated-list-be-converted-cells-column-lt
    
    
    
    
    
    