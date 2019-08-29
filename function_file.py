import pandas
from xlrd.biffh import XLRDError

def open_excel(fileName, sheetName):
    try:
        dataFrame = pandas.read_excel(fileName, sheet_name=sheetName)
        return dataFrame
    except XLRDError as e:
        print('\t' + str(e))
        return None


def open_input_file(filename, sheetName, returnTitle = False):
    dataframe = open_excel(filename,sheetName)
    if dataframe is not None:
        unit = get_unit(dataframe)
        if returnTitle:
            title = dataframe.columns[0]
            return title, unit, dataframe.loc[4:]
        else:
            return unit, dataframe.loc[4:]
    return (None, None, None) if returnTitle else (None, None)

def open_output_file(filename, sheetName):
    return open_excel(filename,sheetName)

def get_unit(datafarme):
    # （単位：件）
    aRow = datafarme.loc[0, :].tolist()
    if '件' in aRow[-1]:
        return '件'
    elif 'kW' in aRow[-1]:
        return 'kW'
