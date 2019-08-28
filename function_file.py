import pandas
from xlrd.biffh import XLRDError

def open_excel(fileName, sheetName):
    try:
        dataFrame = pandas.read_excel(fileName, sheet_name=sheetName)
        return dataFrame
    except XLRDError as e:
        print(e)
        return None


def open_input_file(filename, sheetName):
    df = open_excel(filename,sheetName)
    if df is not None:
        return df.loc[4:]
    return None

def open_output_file(filename, sheetName):
    return open_excel(filename,sheetName)