import function_general
import function_input
import function_excel

# test file reader
fileList = function_general.list_file("input_file")
for aFile in fileList:
    print(aFile, function_general.re_get_date(aFile))
    break

# test input file checker
print(function_input.is_output_file_available("output_file"))

df = function_excel.open_excel("input_file/A_pref201404.xls", function_input.get_sheet_name(0))
print(df)

df = function_excel.open_excel("output_file/Latest 201908 example.xlsx", function_input.get_sheet_name(0))
print(df)