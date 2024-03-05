import pandas as pd
import os

# path to dataframes
folder ="Data Frames\Vytology Oct'23 - Dfs"
dataframes = os.listdir(folder)

write_path = "Extracted\Camelot"

# specify the workbook name
workbook_name = "Vytalogy Financial Results v4_Oct_OCR.xlsx"

final_path = os.path.join(write_path,workbook_name)

with pd.ExcelWriter(final_path) as writer:
    for dataframe in dataframes:
        sheet = dataframe.split(".")[0]
        path = os.path.join(folder,dataframe)
        df = pd.read_excel(path, index_col=0)
        df.to_excel(writer, sheet_name=sheet)  