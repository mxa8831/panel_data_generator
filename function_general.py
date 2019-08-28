import re
import glob

def re_get_date(filename):
    result = re.search("A_pref(20\d\d)(\d\d)", filename)
    if result and len(result.groups()) == 2:
        return (result.group((1)), result.group((2)))
    raise ValueError("The file name did not contain the pattern of year and month as YYYYMM. Example of working pattern: '201903'.")

def list_file(pathname):
    return glob.glob("{}/*.xls*".format(pathname))