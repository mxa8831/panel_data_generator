import function_general
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
]

def get_summary_header_name():
    return summaryHeaderNameList

def get_net_header_name():
    return netHeaderNameList

def is_output_file_available(pathName):
    return len(function_general.list_file(pathName)) != 0

def generate_output(summaryDf, netDf, filename = "output_file/beta_{}.xlsx".format(str(datetime.datetime.now()).replace(":", "."))):
    excelWriter = pandas.ExcelWriter(filename, engine='xlsxwriter')
    summaryDf.to_excel(excelWriter, index=False, sheet_name="Summary1")
    netDf.to_excel(excelWriter, index=False, sheet_name="net")
    excelWriter.save()

