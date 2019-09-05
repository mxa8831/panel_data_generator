import function_general
import function_file
import datetime
import pandas

summaryHeaderNameList = [
    'Year',
    'Month',
    'Prefecture',
    'Sector',
    'Threshold1',
    'Threshold2',
    'Value',
    'Unit',
    'Tab',
    'Title'
]

netHeaderNameList = [
    'Year',
    'Month',
    'Prefecture',
    'Sector',
    'Threshold1',
    'Threshold2',
    'Tab',
    'Title',
    'Net Value',
    'Net Unit',
    'Abnormal'
]

outputSheetNameList = [
    'Summary1',
    'net',
]

def get_summary_header_name():
    return summaryHeaderNameList

def get_net_header_name():
    return netHeaderNameList

def get_output_sheet_name(index):
    if 0 <= index <= 1:
        return outputSheetNameList[index]
    return None

def is_output_file_available(pathName, suffix = '*'):
    return len(function_general.list_file(pathName, suffix)) != 0

def get_latest_output_file(pathName, suffix='*'):
    list_file = function_general.list_file(pathName, suffix)
    if list_file:
        return list_file[-1]
    return None

def generate_output(summaryDf, netDf, filename = "output_file/beta_{}.xlsx".format(str(datetime.datetime.now()).replace(":", "."))):
    excelWriter = pandas.ExcelWriter(filename, engine='xlsxwriter')
    summaryDf.to_excel(excelWriter, index=False, sheet_name="Summary1")
    netDf.to_excel(excelWriter, index=False, sheet_name="net")
    excelWriter.save()

def get_previous_output(finalSummaryDf, finalNetDf):
    previous_filename = get_latest_output_file("output_file", "merged_*")
    print(previous_filename)
    dfSummary = function_file.open_output_file(previous_filename, get_output_sheet_name(0))
    dfNet = function_file.open_output_file(previous_filename, get_output_sheet_name(1))

    if dfSummary is not None: finalSummaryDf = finalSummaryDf.append(dfSummary)
    if dfNet is not None: finalNetDf = finalNetDf.append(dfNet)

    return (finalSummaryDf, finalNetDf)

