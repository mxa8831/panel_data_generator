import function_general
import function_input
import function_output
import function_file
import pandas
import datetime

# finalDataFrame = pandas.DataFrame([], columns=function_output.get_summary_header_name())
finalSummaryData = []
finalNetData = []

for aFile in function_general.list_file("input_file"):
    print("Processing {}".format(aFile))
    year, month = function_general.regex_get_date(aFile)

    title, unit1, df1 = function_file.open_input_file(aFile, function_input.get_sheet_name(0), returnTitle=True)
    unit2, df2 = function_file.open_input_file(aFile, function_input.get_sheet_name(1))
    unit3, df3 = function_file.open_input_file(aFile, function_input.get_sheet_name(2))
    unit4, df4 = function_file.open_input_file(aFile, function_input.get_sheet_name(3))

    if df1 is not None:
        rowlist1 = function_input.process_dataframe(df1, year, month, function_input.get_sheet_name(0),  title, unit1)
        finalSummaryData += [x.to_summary_list() for x in rowlist1]
        finalNetData += [x.to_net_list() for x in rowlist1]
    if df2 is not None:
        rowlist2 = function_input.process_dataframe(df2, year, month, function_input.get_sheet_name(1), title, unit2)
        finalSummaryData += [x.to_summary_list() for x in rowlist2]
        finalNetData += [x.to_net_list() for x in rowlist2]
    if df3 is not None:
        rowlist3 = function_input.process_dataframe(df3, year, month, function_input.get_sheet_name(2), title, unit3)
        finalSummaryData += [x.to_summary_list() for x in rowlist3]
        finalNetData += [x.to_net_list() for x in rowlist3]
    if df4 is not None:
        rowlist4 = function_input.process_dataframe(df4, year, month, function_input.get_sheet_name(3), title, unit4)
        finalSummaryData += [x.to_summary_list() for x in rowlist4]
        finalNetData += [x.to_net_list() for x in rowlist4]

    # set current dataframe to be used next sheet
    function_input.save_current_month_data(year, month, tab1=df1, tab2=df2, tab3=df3, tab4=df4)

print("Generating Output File (xlsx). Will take some time depending on how large the data is...")
newDf = pandas.DataFrame(finalSummaryData, columns=function_output.get_summary_header_name())
netDf = pandas.DataFrame(finalNetData, columns=function_output.get_net_header_name())



filename = "output_file/beta_{}.xlsx".format(str(datetime.datetime.now()).replace(":", "."))
excelWriter = pandas.ExcelWriter(filename, engine='xlsxwriter')
newDf.to_excel(excelWriter, index=False, sheet_name="Summary1")
netDf.to_excel(excelWriter, index=False, sheet_name="net")
excelWriter.save()
