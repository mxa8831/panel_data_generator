import function_general
import function_input
import function_file
import pandas
import datetime

finalData = []

for aFile in function_general.list_file("input_file"):
    print("Processing {}".format(aFile))
    year, month = function_general.re_get_date(aFile)

    df1 = function_file.open_input_file(aFile, function_input.get_sheet_name(0))

    # todo: different processor for different tabs
    # df2 = function_file.open_input_file(aFile, function_input.get_sheet_name(1))
    # df3 = function_file.open_input_file(aFile, function_input.get_sheet_name(2))
    # df4 = function_file.open_input_file(aFile, function_input.get_sheet_name(3))

    if df1 is not None:
        rowlist1 = function_input.process_dataframe(df1, year, month, function_input.get_sheet_name(0))
        finalData += [x.toList() for x in rowlist1]
    # if df2 is not None:
    #     rowlist2 = function_input.process_dataframe(df2, year, month, function_input.get_sheet_name(1))
    #     finalData += [x.toList() for x in rowlist2]
    # if df3 is not None:
    #     rowlist3 = function_input.process_dataframe(df3, year, month, function_input.get_sheet_name(2))
    #     finalData += [x.toList() for x in rowlist3]
    # if df4 is not None:
    #     rowlist4 = function_input.process_dataframe(df4, year, month, function_input.get_sheet_name(3))
    #     finalData += [x.toList() for x in rowlist4]

newDf = pandas.DataFrame(finalData, columns=[
    'Year',
    'Month',
    'Prefecture',
    'Sector',
    'Threshold1',
    'Threshold2',
    'kW'
])
newDf.to_excel("output_file/beta_{}.xlsx".format(str(datetime.datetime.now())), index=False)