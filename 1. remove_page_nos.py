import pandas as pd
import os
import re

# path to dataframes
folder ="Data Frames\Vytology Oct'23 - Dfs"
dataframes = os.listdir(folder)

# remove page numbers
pattern =  re.compile(r"Page \d+ of \d+")

for dataframe in dataframes:
    path = os.path.join(folder,dataframe) 
    df = pd.read_excel(path, index_col=None)
    last_row = df.index[-1]
    data = df.iloc[last_row, :].astype(str)

    if any(pattern.search(cell) for cell in data):
        df= df.drop(last_row)
    
    with pd.ExcelWriter(path) as writer:
        df = df.drop(df.columns[0],axis=1)
        df.to_excel(writer)  
