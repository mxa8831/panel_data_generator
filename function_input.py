import function_general
import pandas

sheetNameList = [
    '表A①－１',
    '表A①－２',
    '表A②－１',
    '表A②－２',
    '表A③',
    '表A④',
]

class Row(object):
    def __init__(self, year, month, perfecture, sector, threshold1, threshold2, value, unit, tab, title):
        self._year = year
        self._month = month
        self._perfecture = perfecture
        self._sector = sector
        self._threshold1 = threshold1
        self._threshold2 = threshold2
        self._value = value
        self._unit = unit
        self._tab = tab
        self._title = title


    def toList(self):
        return [
            self.year,
            self.month,
            self.perfecture,
            self.sector,
            self.threshold1,
            self.threshold2,
            self.value,
            self.unit,
            self.tab,
            self.title
        ]

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
    def value(self):
        return self._value

    @value.setter
    def value(self, value):
        self._value= value

    @property
    def unit(self):
        return self._unit

    @unit.setter
    def unit(self, value):
        self._unit= value

    @property
    def tab(self):
        return self._tab

    @tab.setter
    def tab(self, value):
        self._tab= value

    @property
    def title(self):
        return self._title

    @title.setter
    def title(self, value):
        self._title = value



def get_sheet_name(index):
    if 0 <= index <= 5:
        return sheetNameList[index]
    return None

def is_output_file_available(pathName):
    return len(function_general.list_file(pathName)) != 0

def process_dataframe(dataframe, year, month, tab, title, unit):
    rowList = []
    for index, row in dataframe.iterrows():
        rowList = rowList + process_row(index, row, year, month, tab, title, unit)
    return rowList

def process_row(index, row, year, month, tab, title, unit):
    # print(index, [x for x in row])
    rowList = []

    if year == 2019:
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW未満', '', row[1], unit, tab, title))
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW未満', 'うち自家発電設備併設', row[2], unit, tab, title))
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW以上', '', row[3], unit, tab, title))
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW以上', 'うち50kW未満', row[4], unit, tab, title))
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW以上', 'うち50kW以上500kW未満', row[5], unit, tab, title))
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW以上', 'うち500kW以上1,000kW未満', row[6], unit, tab, title))
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW以上', 'うち1,000kW以上2,000kW未満', row[7], unit, tab, title))
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW以上', 'うち2,000kW以上', row[8], unit, tab, title))

        rowList.append(Row(year, month, row[0], '風力発電設備', '20kW未満', '', row[9], unit, tab, title))
        rowList.append(Row(year, month, row[0], '風力発電設備', '20kW以上', '', row[10], unit, tab, title))
        rowList.append(Row(year, month, row[0], '風力発電設備', '20kW以上', 'うち洋上風力', row[11], unit, tab, title))

        rowList.append(Row(year, month, row[0], '水力発電設備', '200kW未満', '', row[12], unit, tab, title))
        rowList.append(Row(year, month, row[0], '水力発電設備', '200kW未満', 'うち特定水力', row[13], unit, tab, title))
        rowList.append(Row(year, month, row[0], '水力発電設備', '200kW以上 1,000kW未満', '', row[14], unit, tab, title))
        rowList.append(Row(year, month, row[0], '水力発電設備', '200kW以上 1,000kW未満', 'うち特定水力', row[15], unit, tab, title))
        rowList.append(Row(year, month, row[0], '水力発電設備', '1,000kW以上 5,000kW未満', '', row[16], unit, tab, title))
        rowList.append(Row(year, month, row[0], '水力発電設備', '1,000kW以上 5,000kW未満', 'うち特定水力', row[17], unit, tab, title))
        rowList.append(Row(year, month, row[0], '水力発電設備', '5,000kW以上 30,000kW未満', '', row[18], unit, tab, title))
        rowList.append(Row(year, month, row[0], '水力発電設備', '5,000kW以上 30,000kW未満', 'うち特定水力', row[19], unit, tab, title))

        rowList.append(Row(year, month, row[0], '地熱発電設備', '15,000kW未満', '', row[20], unit, tab, title))
        rowList.append(Row(year, month, row[0], '地熱発電設備', '15,000kW以上', '', row[21], unit, tab, title))

        rowList.append(Row(year, month, row[0], 'バイオマス発電設備', 'メタン発酵ガス', '', row[22], unit, tab, title))
        rowList.append(Row(year, month, row[0], 'バイオマス発電設備', '未利用木質', '2,000kW未満', row[23], unit, tab, title))
        rowList.append(Row(year, month, row[0], 'バイオマス発電設備', '未利用木質', '2,000kW以上', row[24], unit, tab, title))
        rowList.append(Row(year, month, row[0], 'バイオマス発電設備', '一般木質・農作物残さ', '', row[25], unit, tab, title))
        rowList.append(Row(year, month, row[0], 'バイオマス発電設備', '建設廃材', '', row[26], unit, tab, title))
        rowList.append(Row(year, month, row[0], 'バイオマス発電設備', '建設廃材', '', row[27], unit, tab, title))

        rowList.append(Row(year, month, row[0], '合計', '', '', row[28], unit, tab, title))

    else:
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW未満', '', row[1], unit, tab, title))
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW未満', 'うち自家発電設備併設', row[2], unit, tab, title))
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW以上', '', row[3], unit, tab, title))
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW以上', 'うち50kW未満', row[4], unit, tab, title))
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW以上', 'うち50kW以上500kW未満', row[5], unit, tab, title))
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW以上', 'うち500kW以上1,000kW未満', row[6], unit, tab, title))
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW以上', 'うち1,000kW以上2,000kW未満', row[7], unit, tab, title))
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW以上', 'うち2,000kW以上', row[8], unit, tab, title))

        rowList.append(Row(year, month, row[0], '風力発電設備', '20kW未満', '', row[9], unit, tab, title))
        rowList.append(Row(year, month, row[0], '風力発電設備', '20kW以上', '', row[10], unit, tab, title))
        rowList.append(Row(year, month, row[0], '風力発電設備', '20kW以上', 'うち洋上風力', row[11], unit, tab, title))

        rowList.append(Row(year, month, row[0], '水力発電設備', '200kW未満', '', row[12], unit, tab, title))
        rowList.append(Row(year, month, row[0], '水力発電設備', '200kW未満', 'うち特定水力', row[13], unit, tab, title))
        rowList.append(Row(year, month, row[0], '水力発電設備', '200kW以上 1,000kW未満', '', row[14], unit, tab, title))
        rowList.append(Row(year, month, row[0], '水力発電設備', '200kW以上 1,000kW未満', 'うち特定水力', row[15], unit, tab, title))
        rowList.append(Row(year, month, row[0], '水力発電設備', '1,000kW以上 5,000kW未満', '', row[16], unit, tab, title))
        rowList.append(Row(year, month, row[0], '水力発電設備', '1,000kW以上 5,000kW未満', 'うち特定水力', row[17], unit, tab, title))

        rowList.append(Row(year, month, row[0], '地熱発電設備', '15,000kW未満', '', row[18], unit, tab, title))
        rowList.append(Row(year, month, row[0], '地熱発電設備', '15,000kW以上', '', row[19], unit, tab, title))

        rowList.append(Row(year, month, row[0], 'バイオマス発電設備', 'メタン発酵ガス', '', row[20], unit, tab, title))
        rowList.append(Row(year, month, row[0], 'バイオマス発電設備', '未利用木質', '', row[21], unit, tab, title))
        rowList.append(Row(year, month, row[0], 'バイオマス発電設備', '一般木質・農作物残さ', '', row[22], unit, tab, title))
        rowList.append(Row(year, month, row[0], 'バイオマス発電設備', '建設廃材', '', row[23], unit, tab, title))
        rowList.append(Row(year, month, row[0], 'バイオマス発電設備', '建設廃材', '', row[24], unit, tab, title))

        rowList.append(Row(year, month, row[0], '合計', '', '', row[25], unit, tab, title))
    # for i in rowList:
    #     print(i.toList())

    return  rowList

