#MacGregor Winegard
#11/5/2021 
# This merges clean output with the 

import pandas as pd

class XL_work:
    def get_xl_sheet(file_path):
        return pd.read_excel(file_path)
    
    
if __name__ == '__main__':
    recip_list = XL_work.get_xl_sheet("XL_Sheets/Advancement Report 10.22.2021.xlsx")