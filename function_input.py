import function_general

sheetNameList = [
    '表A①－１',
    '表A①－２',
    '表A②－１',
    '表A②－２',
    '表A③',
    '表A④',
]

class Row(object):
    def __init__(self, year, month, perfecture, sector, threshold1, threshold2, kw):
        self._year = year
        self._month = month
        self._sector = sector
        self._perfecture = perfecture
        self._threshold1 = threshold1
        self._threshold2 = threshold2
        self._kw = kw

    @property
    def year(self):
        return self._year

    @year.setter
    def year(self, value):
        self._year = value

    @property
    def month(self):
        return self._month

    @month.setter
    def month(self, value):
        self._month= value

    @property
    def perfecture(self):
        return self._perfecture

    @perfecture.setter
    def perfecture(self, value):
        self._perfecture = value

    @property
    def sector(self):
        return self._sector

    @sector.setter
    def sector(self, value):
        self._sector = value

    @property
    def threshold1(self):
        return self._threshold1

    @threshold1.setter
    def threshold1(self, value):
        self._threshold1 = value

    @property
    def threshold2(self):
        return self._threshold2

    @threshold2.setter
    def threshold2(self, value):
        self._threshold2 = value

    @property
    def kw(self):
        return self._kw

    @kw.setter
    def kw(self, value):
        self._kw= value



def get_sheet_name(index):
    if 0 <= index <= 5:
        return sheetNameList[index]
    return None

def is_output_file_available(pathName):
    return len(function_general.list_file(pathName)) != 0


