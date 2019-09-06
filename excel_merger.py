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

    title1, unit1, df1 = function_file.open_input_file(aFile, function_input.get_sheet_name(0), returnTitle=True)
    title2, unit2, df2 = function_file.open_input_file(aFile, function_input.get_sheet_name(1), returnTitle=True)
    title3, unit3, df3 = function_file.open_input_file(aFile, function_input.get_sheet_name(2), returnTitle=True)
    title4, unit4, df4 = function_file.open_input_file(aFile, function_input.get_sheet_name(3), returnTitle=True)

    if df1 is not None:
        rowlist1 = function_input.process_dataframe(df1, year, month, function_input.get_sheet_name(0),  title1, unit1)
        finalSummaryData += [x.to_summary_list() for x in rowlist1]
        finalNetData += [x.to_net_list() for x in rowlist1]
    if df2 is not None:
        rowlist2 = function_input.process_dataframe(df2, year, month, function_input.get_sheet_name(1), title2, unit2)
        finalSummaryData += [x.to_summary_list() for x in rowlist2]
        finalNetData += [x.to_net_list() for x in rowlist2]
    if df3 is not None:
        rowlist3 = function_input.process_dataframe(df3, year, month, function_input.get_sheet_name(2), title3, unit3)
        finalSummaryData += [x.to_summary_list() for x in rowlist3]
        finalNetData += [x.to_net_list() for x in rowlist3]
    if df4 is not None:
        rowlist4 = function_input.process_dataframe(df4, year, month, function_input.get_sheet_name(3), title4, unit4)
        finalSummaryData += [x.to_summary_list() for x in rowlist4]
        finalNetData += [x.to_net_list() for x in rowlist4]

    # set current dataframe to be used next sheet
    function_input.save_current_month_data(year, month, tab1=df1, tab2=df2, tab3=df3, tab4=df4)


# generating dataframe to contains data
finalSummaryDf = pandas.DataFrame([], columns=function_output.get_summary_header_name())
finalNetDf = pandas.DataFrame([], columns=function_output.get_net_header_name())

# attemting to restore previous data (if any)
if function_output.is_output_file_available("output_file", "merged_*"):
    print("\nPrevious Data found. Attempting to merge: ", end='')
    finalSummaryDf, finalNetDf = function_output.get_previous_output(finalSummaryDf, finalNetDf)

print("\nAppending the dataframe")
finalSummaryDf = finalSummaryDf.append(pandas.DataFrame(finalSummaryData, columns=function_output.get_summary_header_name()))
finalNetDf = finalNetDf.append(pandas.DataFrame(finalNetData, columns=function_output.get_net_header_name()))

print("\nGenerating Output File (xlsx). Will take some time depending on how large the data is...")
filename = "output_file/merged_{}.xlsx".format(str(datetime.datetime.now()).replace(":", "."))
excelWriter = pandas.ExcelWriter(filename, engine='xlsxwriter')
finalSummaryDf.to_excel(excelWriter, index=False, sheet_name="Summary1")
finalNetDf.to_excel(excelWriter, index=False, sheet_name="net")
excelWriter.save()
