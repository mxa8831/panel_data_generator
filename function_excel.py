import pandas

def open_excel(fileName, sheetName):
    dataFrame = pandas.read_excel(fileName, sheet_name=sheetName)
    return dataFrame