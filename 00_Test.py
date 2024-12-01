import pandas as pd
import os
from components import *

# Path = os.getcwd() + "/Documents/Excel-Files/weboc_data.xlsx" 
# testPath = os.getcwd() + "/Output.xlsx"

# if os.path.exists(Path):
#     df = pd.read_excel(Path, index_col=None, dtype={'HS Code': str})
#     print(df.head(6))
#     # df.drop(columns=df.columns[0], axis=1, inplace=True)
    
 
# print(df.shape)
# df.drop_duplicates(subset=None, keep='first', inplace=True )
# print(df.shape)   

# df.to_excel("Output.xlsx", index=False, engine='openpyxl')



one = {
    "A": [10, 20, 30],
    "B": "40"
}

two = {
    "A": [40, 60, 90, 60],
    "B": "30"
}

result = {key: one[key] + two[key] for key in one}

print(result)