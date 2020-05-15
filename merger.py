import pandas as pd
from os import listdir
from os import path
from os import walk

for dirpath, dnames, fnames in walk("./input"):
    spreadsheets = []
    index = 0
    for f in fnames:
        if f.endswith(".xlsx"):
            spreadsheet = pd.read_excel(path.join(dirpath, f))
            # drop last row with counts and sums
            spreadsheet.drop(spreadsheet.tail(1).index, inplace=True)
            # only keep headers for first spreadsheet
            if index != 0:
                spreadsheet.drop(spreadsheet.head(6).index, inplace=True)
            spreadsheets.append(spreadsheet)
            index += 1
    print('spreadsheets loaded')
    merged_spreadsheet = pd.concat(spreadsheets)
    merged_spreadsheet.to_excel("merged.xlsx", index=False)
    print('spreadsheets merged')