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
    def __init__(self, year, month, perfecture, sector, threshold1, threshold2, kw):
        self._year = year
        self._month = month
        self._perfecture = perfecture
        self._sector = sector
        self._threshold1 = threshold1
        self._threshold2 = threshold2
        self._kw = kw

    def toList(self):
        return [
            self.year,
            self.month,
            self.perfecture,
            self.sector,
            self.threshold1,
            self.threshold2,
            self.kw,
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

def process_dataframe(dataframe, year, month, tab):
    rowList = []
    for index, row in dataframe.iterrows():
        rowList = rowList + process_row(index, row, year, month)
    return rowList

def process_row(index, row, year, month):
    print(index, [x for x in row])
    rowList = []

    if year == 2019:

        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW未満', '', row[1]))
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW未満', 'うち自家発電設備併設', row[2]))
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW以上', '', row[3]))
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW以上', 'うち50kW未満', row[4]))
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW以上', 'うち50kW以上500kW未満', row[5]))
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW以上', 'うち500kW以上1,000kW未満', row[6]))
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW以上', 'うち1,000kW以上2,000kW未満', row[7]))
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW以上', 'うち2,000kW以上', row[8]))

        rowList.append(Row(year, month, row[0], '風力発電設備', '20kW未満', '', row[9]))
        rowList.append(Row(year, month, row[0], '風力発電設備', '20kW以上', '', row[10]))
        rowList.append(Row(year, month, row[0], '風力発電設備', '20kW以上', 'うち洋上風力', row[11]))

        rowList.append(Row(year, month, row[0], '水力発電設備', '200kW未満', '', row[12]))
        rowList.append(Row(year, month, row[0], '水力発電設備', '200kW未満', 'うち特定水力', row[13]))
        rowList.append(Row(year, month, row[0], '水力発電設備', '200kW以上 1,000kW未満', '', row[14]))
        rowList.append(Row(year, month, row[0], '水力発電設備', '200kW以上 1,000kW未満', 'うち特定水力', row[15]))
        rowList.append(Row(year, month, row[0], '水力発電設備', '1,000kW以上 5,000kW未満', '', row[16]))
        rowList.append(Row(year, month, row[0], '水力発電設備', '1,000kW以上 5,000kW未満', 'うち特定水力', row[17]))
        rowList.append(Row(year, month, row[0], '水力発電設備', '5,000kW以上 30,000kW未満', '', row[18]))
        rowList.append(Row(year, month, row[0], '水力発電設備', '5,000kW以上 30,000kW未満', 'うち特定水力', row[19]))

        rowList.append(Row(year, month, row[0], '地熱発電設備', '15,000kW未満', '', row[20]))
        rowList.append(Row(year, month, row[0], '地熱発電設備', '15,000kW以上', '', row[21]))

        rowList.append(Row(year, month, row[0], 'バイオマス発電設備', 'メタン発酵ガス', '', row[22]))
        rowList.append(Row(year, month, row[0], 'バイオマス発電設備', '未利用木質', '2,000kW未満', row[23]))
        rowList.append(Row(year, month, row[0], 'バイオマス発電設備', '未利用木質', '2,000kW以上', row[24]))
        rowList.append(Row(year, month, row[0], 'バイオマス発電設備', '一般木質・農作物残さ', '', row[25]))
        rowList.append(Row(year, month, row[0], 'バイオマス発電設備', '建設廃材', '', row[26]))
        rowList.append(Row(year, month, row[0], 'バイオマス発電設備', '建設廃材', '', row[27]))

        rowList.append(Row(year, month, row[0], '合計', '', '', row[28]))

    else:
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW未満', '', row[1]))
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW未満', 'うち自家発電設備併設', row[2]))
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW以上', '', row[3]))
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW以上', 'うち50kW未満', row[4]))
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW以上', 'うち50kW以上500kW未満', row[5]))
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW以上', 'うち500kW以上1,000kW未満', row[6]))
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW以上', 'うち1,000kW以上2,000kW未満', row[7]))
        rowList.append(Row(year, month, row[0], '太陽光発電設備', '10kW以上', 'うち2,000kW以上', row[8]))

        rowList.append(Row(year, month, row[0], '風力発電設備', '20kW未満', '', row[9]))
        rowList.append(Row(year, month, row[0], '風力発電設備', '20kW以上', '', row[10]))
        rowList.append(Row(year, month, row[0], '風力発電設備', '20kW以上', 'うち洋上風力', row[11]))

        rowList.append(Row(year, month, row[0], '水力発電設備', '200kW未満', '', row[12]))
        rowList.append(Row(year, month, row[0], '水力発電設備', '200kW未満', 'うち特定水力', row[13]))
        rowList.append(Row(year, month, row[0], '水力発電設備', '200kW以上 1,000kW未満', '', row[14]))
        rowList.append(Row(year, month, row[0], '水力発電設備', '200kW以上 1,000kW未満', 'うち特定水力', row[15]))
        rowList.append(Row(year, month, row[0], '水力発電設備', '1,000kW以上 5,000kW未満', '', row[16]))
        rowList.append(Row(year, month, row[0], '水力発電設備', '1,000kW以上 5,000kW未満', 'うち特定水力', row[17]))

        rowList.append(Row(year, month, row[0], '地熱発電設備', '15,000kW未満', '', row[18]))
        rowList.append(Row(year, month, row[0], '地熱発電設備', '15,000kW以上', '', row[19]))

        rowList.append(Row(year, month, row[0], 'バイオマス発電設備', 'メタン発酵ガス', '', row[20]))
        rowList.append(Row(year, month, row[0], 'バイオマス発電設備', '未利用木質', '', row[21]))
        rowList.append(Row(year, month, row[0], 'バイオマス発電設備', '一般木質・農作物残さ', '', row[22]))
        rowList.append(Row(year, month, row[0], 'バイオマス発電設備', '建設廃材', '', row[23]))
        rowList.append(Row(year, month, row[0], 'バイオマス発電設備', '建設廃材', '', row[24]))

        rowList.append(Row(year, month, row[0], '合計', '', '', row[25]))
    # for i in rowList:
    #     print(i.toList())

    return  rowList

