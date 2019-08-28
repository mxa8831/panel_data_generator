import function_general
import function_input
import function_file
import pandas

# test file reader
fileList = function_general.list_file("input_file")
for aFile in fileList:
    print(aFile, function_general.re_get_date(aFile))
    break

# test input file checker
print(function_input.is_output_file_available("output_file"))

# opening input file
year, month = function_general.re_get_date("input_file/A_pref201903.xlsx")
print(year, month)
df = function_file.open_input_file("input_file/A_pref201903.xlsx", function_input.get_sheet_name(0))
rowlist = function_input.process_dataframe(df, year, month)

newDf = pandas.DataFrame([x.toList() for x in rowlist], columns=[
    'Year',
    'Month',
    'Prefecture',
    'Sector',
    'Threshold1',
    'Threshold2',
    'kW'
])
print(newDf)

# dataframe to list
print(newDf.values.tolist())

# newDf.to_excel("output_file/beta.xlsx", index=False)

# print(df)


# opening output file
# print(function_excel.open_excel("output_file/Latest 201908 example.xlsx", function_input.get_sheet_name(0)))
# print(function_excel.open_excel("output_file/Latest 201908 example.xlsx", function_input.get_sheet_name(1)))
