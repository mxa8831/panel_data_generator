import function_general

sheetNameList = [
    '表A①－１',
    '表A①－２',
    '表A②－１',
    '表A②－２',
    '表A③',
    '表A④',
]

def get_sheet_name(index):
    if 0 <= index <= 5:
        return sheetNameList[index]
    return None

def is_output_file_available(pathName):
    return len(function_general.list_file(pathName)) != 0


