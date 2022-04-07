import os
import pandas as pd
from os import walk

# Choose new merged file's name
NAME_OF_NEW_XLSX_FILE = 'output.xlsx'

# Enter in LIST_OF_COLUMNS all the columns that you want to merge
LIST_OF_COLUMNS = ['try1','try2','try3']

# Put all xlsx files in the same (any) folder, adding MergeExcels.py code file
PATH_TO_EXCEL_SHEETS = os.getcwd()

# Program gets all files names in folder
excel_sheets_names = []

for (dirpath, dirnames, filenames) in walk(PATH_TO_EXCEL_SHEETS):
    excel_sheets_names.extend(filenames)
    break

# Removing all non-excel files and last output file from about-to-be-merged list
for filename in filenames:
    if (not filename.endswith('.xlsx')) or filename. __eq__(NAME_OF_NEW_XLSX_FILE):
        excel_sheets_names.remove(str(filename))

# Printing the files that will be merged
print(len(excel_sheets_names))
print(excel_sheets_names)

data_frame = []
for excel in excel_sheets_names:
    df = pd.read_excel(excel)
    values = df[LIST_OF_COLUMNS]
    data_frame.append(values)

join = pd.concat(data_frame)

join.to_excel(NAME_OF_NEW_XLSX_FILE)
