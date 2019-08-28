import pandas

def open_excel(fileName, sheetName):
    dataFrame = pandas.read_excel(fileName, sheet_name=sheetName)
    return dataFrame

def open_input_file(filename, sheetName):
    df = open_excel(filename,sheetName)
    return df.loc[4:]

def open_output_file(filename, sheetName):
    return open_excel(filename,sheetName)