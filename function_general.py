import re
import glob

def regex_get_date(filename):
    result = re.search("A_pref(20\d\d)(\d\d)", filename)
    if result and len(result.groups()) == 2:
        return (result.group((1)), result.group((2)))
    raise ValueError("The file name did not contain the pattern of year and month as YYYYMM. Example of working pattern: '201903'.")

def list_file(pathname):
    return sorted(glob.glob("{}/*.xls*".format(pathname)))

def string_to_int(string):
    return int(string)

def get_month_before(year, month):
    year = string_to_int(year)
    month = string_to_int(month)
    if month == 1:
        return (year - 1, 12)
    return (year, month - 1)

def compute_difference(currentValue, previousValue):
    if previousValue is None or previousValue is 'NA':
        return 'NA'
    return string_to_int(currentValue) - string_to_int(previousValue)