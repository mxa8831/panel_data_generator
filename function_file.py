import pandas
from xlrd.biffh import XLRDError

def open_excel(fileName, sheetName):
    try:
        dataFrame = pandas.read_excel(fileName, sheet_name=sheetName)
        return dataFrame
    except XLRDError as e:
        print('\t' + str(e))
        return None


def open_input_file(filename, sheetName, returnTitleandUnit = False):
    dataframe = open_excel(filename,sheetName)
    if dataframe is not None:
        if returnTitleandUnit:
            title = dataframe.columns[0]
            unit = get_unit(dataframe)
            return title, unit, dataframe.loc[4:]
        else:
            return dataframe.loc[4:]
    return (None, None, None) if returnTitleandUnit else None

def open_output_file(filename, sheetName):
    return open_excel(filename,sheetName)

def get_unit(datafarme):
    # （単位：件）
    aRow = datafarme.loc[0, :].tolist()
    if '件' in aRow[-1]:
        return '件'
    elif 'kW' in aRow[-1]:
        return 'kW'
