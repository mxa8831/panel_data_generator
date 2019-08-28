import function_general
import function_input

# filename = 'input_file/A_pref201404.xls'
# dateYear, dateMonth = function_general.re_get_date(filename)
# print(dateYear, dateMonth)


fileList = function_general.list_file("input_file/*.xls")

for aFile in fileList:
    print(aFile, function_general.re_get_date(aFile))
    break