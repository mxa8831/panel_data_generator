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
    dataframe = open_excel(filename,sheetName)
    if dataframe is not None:
        title = dataframe.columns[0]
        unit = get_unit(dataframe)
        return title, unit, dataframe.loc[4:]
    return (None, None, None)

def open_output_file(filename, sheetName):
    return open_excel(filename,sheetName)

def get_unit(datafarme):
    # （単位：件）
    aRow = datafarme.loc[0, :].tolist()
    return aRow[-1]
