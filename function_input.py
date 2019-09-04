import function_general
from function_general import compute_difference as cd

sheetNameList = [
    '表A①－１',
    '表A①－２',
    '表A②－１',
    '表A②－２',
    '表A③',
    '表A④',
]

class Row(object):
    def __init__(self, year, month, perfecture, sector, threshold1, threshold2, value, unit, tab, title, difference = 'NA'):
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
        self._difference = difference
        self._abnormal = True if difference is 'NA' else False


    def to_summary_list(self):
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

    def to_net_list(self):
        return [
            self.year,
            self.month,
            self.perfecture,
            self.sector,
            self.threshold1,
            self.threshold2,
            self.tab,
            self.title,
            self.difference,
            self.unit,
            ('Yes' if self.abnormal else 'No')
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

    @property
    def difference(self):
        return self._difference

    @difference.setter
    def difference(self, value):
        self._difference = value

    @property
    def abnormal(self):
        return self._abnormal

    @abnormal.setter
    def abnormal(self, value):
        self._abnormal= value


def get_sheet_name(index):
    if 0 <= index <= 5:
        return sheetNameList[index]
    return None

last_month_data = {
    'year': 2019,
    'month': 2,
    '表A①－１': None, # df tab 1
    '表A①－２': None, # df tab 2
    '表A②－１': None, # df tab 3
    '表A②－２': None, # df tab 4
}

def save_current_month_data(year, month, tab1 = None, tab2 = None, tab3 = None, tab4 = None):
    global last_month_data
    last_month_data = {
        'year': function_general.string_to_int(year),
        'month': function_general.string_to_int(month),
        '表A①－１': tab1, # df tab 1
        '表A①－２': tab2, # df tab 2
        '表A②－１': tab3, # df tab 3
        '表A②－２': tab4, # df tab 4
    }

def get_previous_month_data(currentYear, currentMonth, tab):
    (monthBeforeYear, monthBeforemonth) = function_general.get_month_before(currentYear, currentMonth)
    global last_month_data
    if last_month_data['year'] == function_general.string_to_int(monthBeforeYear) and \
            last_month_data['month'] == function_general.string_to_int(monthBeforemonth) and \
            last_month_data[tab] is not None:
        return last_month_data[tab]
    return None

def process_dataframe(currentDataframe, currentYear, currentMonth, tab, title, unit):
    rowList = []
    previousDataframe = get_previous_month_data(currentYear, currentMonth, tab)

    if previousDataframe is not None:
        # Last month's data is there, so we do coupled iteration with zip
        for (currentDataframeIndex, currentDataframeRow), (prevoiusDataframeIndex, previousDataframeRow) \
                in zip(currentDataframe.iterrows(), previousDataframe.iterrows()):
            if tab == get_sheet_name(0):
                rowList = rowList + process_row_tab_1(currentDataframeIndex, currentDataframeRow, currentYear, currentMonth, tab, title, unit, previousDataframeRow)
            elif tab == get_sheet_name(1):
                rowList = rowList + process_row_tab_2(currentDataframeIndex, currentDataframeRow, currentYear, currentMonth, tab, title, unit, previousDataframeRow)
            elif tab == get_sheet_name(2):
                rowList = rowList + process_row_tab_3(currentDataframeIndex, currentDataframeRow, currentYear, currentMonth, tab, title, unit, previousDataframeRow)
            elif tab == get_sheet_name(3):
                rowList = rowList + process_row_tab_4(currentDataframeIndex, currentDataframeRow, currentYear, currentMonth, tab, title, unit, previousDataframeRow)
    else:
        # Last month's data is not there, so we do normal iteration
        for currentDataframeIndex, currentDataframeRow in currentDataframe.iterrows():
            if tab == get_sheet_name(0):
                rowList = rowList + process_row_tab_1(currentDataframeIndex, currentDataframeRow, currentYear, currentMonth, tab, title, unit)
            elif tab == get_sheet_name(1):
                rowList = rowList + process_row_tab_2(currentDataframeIndex, currentDataframeRow, currentYear, currentMonth, tab, title, unit)
            elif tab == get_sheet_name(2):
                rowList = rowList + process_row_tab_3(currentDataframeIndex, currentDataframeRow, currentYear, currentMonth, tab, title, unit)
            elif tab == get_sheet_name(3):
                rowList = rowList + process_row_tab_4(currentDataframeIndex, currentDataframeRow, currentYear, currentMonth, tab, title, unit)

    return rowList

def process_row_tab_1(index, currentMonthRow, year, month, tab, title, unit, previousMonthRow = None):
    # print(index, [x for x in row])
    rowList = []

    if year == 2019:
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW未満', '', currentMonthRow[1], unit, tab, title, cd(1, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW未満', 'うち自家発電設備併設', currentMonthRow[2], unit, tab, title, cd(2, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', '', currentMonthRow[3], unit, tab, title, cd(3, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち50kW未満', currentMonthRow[4], unit, tab, title, cd(4, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち50kW以上500kW未満', currentMonthRow[5], unit, tab, title, cd(5, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち500kW以上1,000kW未満', currentMonthRow[6], unit, tab, title, cd(6, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち1,000kW以上2,000kW未満', currentMonthRow[7], unit, tab, title, cd(7, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち2,000kW以上', currentMonthRow[8], unit, tab, title, cd(8, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '風力発電設備', '20kW未満', '', currentMonthRow[9], unit, tab, title, cd(9, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '風力発電設備', '20kW以上', '', currentMonthRow[10], unit, tab, title, cd(10, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '風力発電設備', '20kW以上', 'うち洋上風力', currentMonthRow[11], unit, tab, title, cd(11, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW未満', '', currentMonthRow[12], unit, tab, title, cd(12, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW未満', 'うち特定水力', currentMonthRow[13], unit, tab, title, cd(13, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW以上 1,000kW未満', '', currentMonthRow[14], unit, tab, title, cd(14, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW以上 1,000kW未満', 'うち特定水力', currentMonthRow[15], unit, tab, title, cd(15, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '1,000kW以上 5,000kW未満', '', currentMonthRow[16], unit, tab, title, cd(16, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '1,000kW以上 5,000kW未満', 'うち特定水力', currentMonthRow[17], unit, tab, title, cd(17, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '5,000kW以上 30,000kW未満', '', currentMonthRow[18], unit, tab, title, cd(18, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '5,000kW以上 30,000kW未満', 'うち特定水力', currentMonthRow[19], unit, tab, title, cd(19, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '地熱発電設備', '15,000kW未満', '', currentMonthRow[20], unit, tab, title, cd(20, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '地熱発電設備', '15,000kW以上', '', currentMonthRow[21], unit, tab, title, cd(20, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備', 'メタン発酵ガス', '', currentMonthRow[22], unit, tab, title, cd(22, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備', '未利用木質', '2,000kW未満', currentMonthRow[23], unit, tab, title, cd(23, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備', '未利用木質', '2,000kW以上', currentMonthRow[24], unit, tab, title, cd(24, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備', '一般木質・農作物残さ', '', currentMonthRow[25], unit, tab, title, cd(25, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備', '建設廃材', '', currentMonthRow[26], unit, tab, title, cd(26, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備', '一般廃棄物・木質以外', '', currentMonthRow[27], unit, tab, title, cd(27, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '合計', '', '', currentMonthRow[28], unit, tab, title, cd(28, currentMonthRow, previousMonthRow)))

    else:
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW未満', '', currentMonthRow[1], unit, tab, title, cd(1, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW未満', 'うち自家発電設備併設', currentMonthRow[2], unit, tab, title, cd(2, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', '', currentMonthRow[3], unit, tab, title, cd(3, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち50kW未満', currentMonthRow[4], unit, tab, title, cd(4, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち50kW以上500kW未満', currentMonthRow[5], unit, tab, title, cd(5, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち500kW以上1,000kW未満', currentMonthRow[6], unit, tab, title, cd(6, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち1,000kW以上2,000kW未満', currentMonthRow[7], unit, tab, title, cd(7, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち2,000kW以上', currentMonthRow[8], unit, tab, title, cd(8, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '風力発電設備', '20kW未満', '', currentMonthRow[9], unit, tab, title, cd(9, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '風力発電設備', '20kW以上', '', currentMonthRow[10], unit, tab, title, cd(10, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '風力発電設備', '20kW以上', 'うち洋上風力', currentMonthRow[11], unit, tab, title, cd(11, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW未満', '', currentMonthRow[12], unit, tab, title, cd(12, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW未満', 'うち特定水力', currentMonthRow[13], unit, tab, title, cd(13, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW以上 1,000kW未満', '', currentMonthRow[14], unit, tab, title, cd(14, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW以上 1,000kW未満', 'うち特定水力', currentMonthRow[15], unit, tab, title, cd(15, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '1,000kW以上 5,000kW未満', '', currentMonthRow[16], unit, tab, title, cd(16, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '1,000kW以上 5,000kW未満', 'うち特定水力', currentMonthRow[17], unit, tab, title, cd(17, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '地熱発電設備', '15,000kW未満', '', currentMonthRow[18], unit, tab, title, cd(18, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '地熱発電設備', '15,000kW以上', '', currentMonthRow[19], unit, tab, title, cd(19, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備', 'メタン発酵ガス', '', currentMonthRow[20], unit, tab, title, cd(20, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備', '未利用木質', '', currentMonthRow[21], unit, tab, title, cd(21, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備', '一般木質・農作物残さ', '', currentMonthRow[22], unit, tab, title, cd(22, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備', '建設廃材', '', currentMonthRow[23], unit, tab, title, cd(23, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備', '一般廃棄物・木質以外', '', currentMonthRow[24], unit, tab, title, cd(24, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '合計', '', '', currentMonthRow[25], unit, tab, title, cd(25, currentMonthRow, previousMonthRow)))
    return  rowList

def process_row_tab_2(index, currentMonthRow, year, month, tab, title, unit, previousMonthRow = None):
    # print(index, [x for x in row])
    rowList = []

    if year == 2019:
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW未満', '', currentMonthRow[1], unit, tab, title, cd(1, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', '', currentMonthRow[2], unit, tab, title, cd(2, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち50kW未満', currentMonthRow[3], unit, tab, title, cd(3, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち50kW以上500kW未満', currentMonthRow[4], unit, tab, title, cd(4, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち500kW以上1,000kW未満', currentMonthRow[5], unit, tab, title, cd(5, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち1,000kW以上2,000kW未満', currentMonthRow[6], unit, tab, title, cd(6, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち2,000kW以上', currentMonthRow[7], unit, tab, title, cd(7, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '風力発電設備', '20kW未満', '', currentMonthRow[8], unit, tab, title, cd(8, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '風力発電設備', '20kW以上', '', currentMonthRow[9], unit, tab, title, cd(9, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '風力発電設備', '20kW以上', 'うち洋上風力', currentMonthRow[10], unit, tab, title, cd(10, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW未満', '', currentMonthRow[11], unit, tab, title, cd(11, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW未満', 'うち特定水力', currentMonthRow[12], unit, tab, title, cd(12, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW以上 1,000kW未満', '', currentMonthRow[13], unit, tab, title, cd(13, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW以上 1,000kW未満', 'うち特定水力', currentMonthRow[14], unit, tab, title, cd(14, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '1,000kW以上 5,000kW未満', '', currentMonthRow[15], unit, tab, title, cd(15, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '1,000kW以上 5,000kW未満', 'うち特定水力', currentMonthRow[16], unit, tab, title, cd(16, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '5,000kW以上 30,000kW未満', '', currentMonthRow[17], unit, tab, title, cd(17, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '5,000kW以上 30,000kW未満', 'うち特定水力', currentMonthRow[18], unit, tab, title, cd(18, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '地熱発電設備', '15,000kW未満', '', currentMonthRow[19], unit, tab, title, cd(19, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '地熱発電設備', '15,000kW以上', '', currentMonthRow[20], unit, tab, title, cd(20, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備', 'メタン発酵ガス', '', currentMonthRow[21], unit, tab, title, cd(21, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備', '未利用木質', '2,000kW未満', currentMonthRow[22], unit, tab, title, cd(22, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備', '未利用木質', '2,000kW以上', currentMonthRow[23], unit, tab, title, cd(23, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備', '一般木質・農作物残さ', '', currentMonthRow[24], unit, tab, title, cd(24, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備', '建設廃材', '', currentMonthRow[25], unit, tab, title, cd(25, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備', '一般廃棄物・木質以外', '', currentMonthRow[26], unit, tab, title, cd(26, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '合計', '', '', currentMonthRow[27], unit, tab, title, cd(27, currentMonthRow, previousMonthRow)))
    else:
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW未満', '', currentMonthRow[1], unit, tab, title, cd(1, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', '', currentMonthRow[2], unit, tab, title, cd(2, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち50kW未満', currentMonthRow[3], unit, tab, title, cd(3, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち50kW以上500kW未満', currentMonthRow[4], unit, tab, title, cd(4, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち500kW以上1,000kW未満', currentMonthRow[5], unit, tab, title, cd(5, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち1,000kW以上2,000kW未満', currentMonthRow[6], unit, tab, title, cd(6, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち2,000kW以上', currentMonthRow[7], unit, tab, title, cd(7, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '風力発電設備', '20kW未満', '', currentMonthRow[8], unit, tab, title, cd(8, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '風力発電設備', '20kW以上', '', currentMonthRow[9], unit, tab, title, cd(9, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '風力発電設備', '20kW以上', 'うち洋上風力', currentMonthRow[10], unit, tab, title, cd(10, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW未満', '', currentMonthRow[11], unit, tab, title, cd(11, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW未満', 'うち特定水力', currentMonthRow[12], unit, tab, title, cd(12, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW以上 1,000kW未満', '', currentMonthRow[13], unit, tab, title, cd(13, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW以上 1,000kW未満', 'うち特定水力', currentMonthRow[14], unit, tab, title, cd(14, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '1,000kW以上 30,000kW未満', '', currentMonthRow[15], unit, tab, title, cd(15, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '1,000kW以上 30,000kW未満', 'うち特定水力', currentMonthRow[16], unit, tab, title, cd(16, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '地熱発電設備', '15,000kW未満', '', currentMonthRow[17], unit, tab, title, cd(17, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '地熱発電設備', '15,000kW以上', '', currentMonthRow[18], unit, tab, title, cd(18, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備', 'メタン発酵ガス', '', currentMonthRow[19], unit, tab, title, cd(19, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備', '未利用木質', '', currentMonthRow[20], unit, tab, title, cd(20, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備', '一般木質・農作物残さ', '', currentMonthRow[21], unit, tab, title, cd(21, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備', '建設廃材', '', currentMonthRow[22], unit, tab, title, cd(22, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備', '一般廃棄物・木質以外', '', currentMonthRow[23], unit, tab, title, cd(23, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '合計', '', '', currentMonthRow[24], unit, tab, title, cd(24, currentMonthRow, previousMonthRow)))
    return rowList

def process_row_tab_3(index, currentMonthRow, year, month, tab, title, unit, previousMonthRow = None):
    # print(index, [x for x in row])
    rowList = []

    if year == 2019:
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW未満', '', currentMonthRow[1], unit, tab, title, cd(1, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', '', currentMonthRow[2], unit, tab, title, cd(2, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち50kW未満', currentMonthRow[3], unit, tab, title, cd(3, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち50kW以上500kW未満', currentMonthRow[4], unit, tab, title, cd(4, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち500kW以上1,000kW未満', currentMonthRow[5], unit, tab, title, cd(5, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち1,000kW以上2,000kW未満', currentMonthRow[6], unit, tab, title, cd(6, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち2,000kW以上', currentMonthRow[7], unit, tab, title, cd(7, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '風力発電設備', '20kW未満', '', currentMonthRow[8], unit, tab, title, cd(8, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '風力発電設備', '20kW以上', '', currentMonthRow[9], unit, tab, title, cd(9, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '風力発電設備', '20kW以上', 'うち洋上風力', currentMonthRow[10], unit, tab, title, cd(10, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW未満', '', currentMonthRow[11], unit, tab, title, cd(11, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW未満', 'うち特定水力', currentMonthRow[12], unit, tab, title, cd(12, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW以上 1,000kW未満', '', currentMonthRow[13], unit, tab, title, cd(13, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW以上 1,000kW未満', 'うち特定水力', currentMonthRow[14], unit, tab, title, cd(14, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '1,000kW以上 5,000kW未満', '', currentMonthRow[15], unit, tab, title, cd(15, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '1,000kW以上 5,000kW未満', 'うち特定水力', currentMonthRow[16], unit, tab, title, cd(16, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '5,000kW以上 30,000kW未満', '', currentMonthRow[17], unit, tab, title, cd(17, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '5,000kW以上 30,000kW未満', 'うち特定水力', currentMonthRow[18], unit, tab, title, cd(18, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '地熱発電設備', '15,000kW未満', '', currentMonthRow[19], unit, tab, title, cd(19, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '地熱発電設備', '15,000kW以上', '', currentMonthRow[20], unit, tab, title, cd(20, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮なし）', 'メタン発酵ガス', '', currentMonthRow[21], unit, tab, title, cd(21, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮なし）', '未利用木質', '2,000kW未満', currentMonthRow[22], unit, tab, title, cd(22, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮なし）', '未利用木質', '2,000kW以上', currentMonthRow[23], unit, tab, title, cd(23, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮なし）', '一般木質・農作物残さ', '', currentMonthRow[24], unit, tab, title, cd(24, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮なし）', '建設廃材', '', currentMonthRow[25], unit, tab, title, cd(25, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮なし）', '一般廃棄物・木質以外', '', currentMonthRow[26], unit, tab, title, cd(26, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮あり）', 'メタン発酵ガス', '', currentMonthRow[27], unit, tab, title, cd(27, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮あり）', '未利用木質', '2,000kW未満', currentMonthRow[28], unit, tab, title, cd(28, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮あり）', '未利用木質', '2,000kW以上', currentMonthRow[29], unit, tab, title, cd(29, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮あり）', '一般木質・農作物残さ', '', currentMonthRow[30], unit, tab, title, cd(30, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮あり）', '建設廃材', '', currentMonthRow[31], unit, tab, title, cd(31, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮あり）', '一般廃棄物・木質以外', '', currentMonthRow[32], unit, tab, title, cd(32, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '合計（バイオマス発電設備については、バイオマス比率を考慮したものを合計）', '', '', currentMonthRow[33], unit, tab, title, cd(33, currentMonthRow, previousMonthRow)))
    else:
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW未満', '', currentMonthRow[1], unit, tab, title, cd(1, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW未満', 'うち自家発電設備併設', currentMonthRow[2], unit, tab, title, cd(2, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', '', currentMonthRow[3], unit, tab, title, cd(3, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち50kW未満', currentMonthRow[4], unit, tab, title, cd(4, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち50kW以上500kW未満', currentMonthRow[5], unit, tab, title, cd(5, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち500kW以上1,000kW未満', currentMonthRow[6], unit, tab, title, cd(6, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち1,000kW以上2,000kW未満', currentMonthRow[7], unit, tab, title, cd(7, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち2,000kW以上', currentMonthRow[8], unit, tab, title, cd(8, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '風力発電設備', '20kW未満', '', currentMonthRow[9], unit, tab, title, cd(9, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '風力発電設備', '20kW以上', '', currentMonthRow[10], unit, tab, title, cd(10, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '風力発電設備', '20kW以上', 'うち洋上風力', currentMonthRow[11], unit, tab, title, cd(11, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW未満', '', currentMonthRow[12], unit, tab, title, cd(12, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW未満', 'うち特定水力', currentMonthRow[13], unit, tab, title, cd(13, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW以上 1,000kW未満', '', currentMonthRow[14], unit, tab, title, cd(14, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW以上 1,000kW未満', 'うち特定水力', currentMonthRow[15], unit, tab, title, cd(15, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '1,000kW以上 30,000kW未満', '', currentMonthRow[16], unit, tab, title, cd(16, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '1,000kW以上 30,000kW未満', 'うち特定水力', currentMonthRow[17], unit, tab, title, cd(17, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '地熱発電設備', '15,000kW未満', '', currentMonthRow[18], unit, tab, title, cd(18, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '地熱発電設備', '15,000kW以上', '', currentMonthRow[19], unit, tab, title, cd(19, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮なし）', 'メタン発酵ガス', '', currentMonthRow[20], unit, tab, title, cd(20, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮なし）', '未利用木質', '', currentMonthRow[21], unit, tab, title, cd(21, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮なし）', '一般木質・農作物残さ', '', currentMonthRow[22], unit, tab, title, cd(22, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮なし）', '建設廃材', '', currentMonthRow[23], unit, tab, title, cd(23, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮なし）', '一般廃棄物・木質以外', '', currentMonthRow[24], unit, tab, title, cd(24, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮あり）', 'メタン発酵ガス', '', currentMonthRow[25], unit, tab, title, cd(25, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮あり）', '未利用木質', '', currentMonthRow[26], unit, tab, title, cd(26, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮あり）', '一般木質・農作物残さ', '', currentMonthRow[27], unit, tab, title, cd(27, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮あり）', '建設廃材', '', currentMonthRow[28], unit, tab, title, cd(28, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮あり）', '一般廃棄物・木質以外', '', currentMonthRow[29], unit, tab, title, cd(29, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '合計（バイオマス発電設備については、バイオマス比率を考慮したものを合計）', '', '', currentMonthRow[30], unit, tab, title, cd(30, currentMonthRow, previousMonthRow)))
    return rowList

def process_row_tab_4(index, currentMonthRow, year, month, tab, title, unit, previousMonthRow = None):
    # print(index, [x for x in row])
    rowList = []

    if year == 2019:
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW未満', '', currentMonthRow[1], unit, tab, title, cd(1, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', '', currentMonthRow[2], unit, tab, title, cd(2, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち50kW未満', currentMonthRow[3], unit, tab, title, cd(3, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち50kW以上500kW未満', currentMonthRow[4], unit, tab, title, cd(4, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち500kW以上1,000kW未満', currentMonthRow[5], unit, tab, title, cd(5, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち1,000kW以上2,000kW未満', currentMonthRow[6], unit, tab, title, cd(6, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち2,000kW以上', currentMonthRow[7], unit, tab, title, cd(7, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '風力発電設備', '20kW未満', '', currentMonthRow[8], unit, tab, title, cd(8, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '風力発電設備', '20kW以上', '', currentMonthRow[9], unit, tab, title, cd(9, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '風力発電設備', '20kW以上', 'うち洋上風力', currentMonthRow[10], unit, tab, title, cd(10, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW未満', '', currentMonthRow[11], unit, tab, title, cd(11, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW未満', 'うち特定水力', currentMonthRow[12], unit, tab, title, cd(12, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW以上 1,000kW未満', '', currentMonthRow[13], unit, tab, title, cd(13, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW以上 1,000kW未満', 'うち特定水力', currentMonthRow[14], unit, tab, title, cd(14, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '1,000kW以上 5,000kW未満', '', currentMonthRow[15], unit, tab, title, cd(15, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '1,000kW以上 5,000kW未満', 'うち特定水力', currentMonthRow[16], unit, tab, title, cd(16, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '5,000kW以上 30,000kW未満', '', currentMonthRow[17], unit, tab, title, cd(17, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '5,000kW以上 30,000kW未満', 'うち特定水力', currentMonthRow[18], unit, tab, title, cd(18, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '地熱発電設備', '15,000kW未満', '', currentMonthRow[19], unit, tab, title, cd(19, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '地熱発電設備', '15,000kW以上', '', currentMonthRow[20], unit, tab, title, cd(20, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮なし）', 'メタン発酵ガス', '', currentMonthRow[21], unit, tab, title, cd(21, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮なし）', '未利用木質', '2,000kW未満', currentMonthRow[22], unit, tab, title, cd(22, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮なし）', '未利用木質', '2,000kW以上', currentMonthRow[23], unit, tab, title, cd(23, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮なし）', '一般木質・農作物残さ', '', currentMonthRow[24], unit, tab, title, cd(24, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮なし）', '建設廃材', '', currentMonthRow[25], unit, tab, title, cd(25, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮なし）', '一般廃棄物・木質以外', '', currentMonthRow[26], unit, tab, title, cd(26, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮あり）', 'メタン発酵ガス', '', currentMonthRow[27], unit, tab, title, cd(27, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮あり）', '未利用木質', '2,000kW未満', currentMonthRow[28], unit, tab, title, cd(28, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮あり）', '未利用木質', '2,000kW以上', currentMonthRow[29], unit, tab, title, cd(29, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮あり）', '一般木質・農作物残さ', '', currentMonthRow[30], unit, tab, title, cd(30, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮あり）', '建設廃材', '', currentMonthRow[31], unit, tab, title, cd(31, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮あり）', '一般廃棄物・木質以外', '', currentMonthRow[32], unit, tab, title, cd(32, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '合計（バイオマス発電設備については、バイオマス比率を考慮したものを合計）', '', '', currentMonthRow[33], unit, tab, title, cd(33, currentMonthRow, previousMonthRow)))
    else:
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW未満', '', currentMonthRow[1], unit, tab, title, cd(1, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', '', currentMonthRow[2], unit, tab, title, cd(2, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち50kW未満', currentMonthRow[3], unit, tab, title, cd(3, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち50kW以上500kW未満', currentMonthRow[4], unit, tab, title, cd(4, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち500kW以上1,000kW未満', currentMonthRow[5], unit, tab, title, cd(5, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち1,000kW以上2,000kW未満', currentMonthRow[6], unit, tab, title, cd(6, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '太陽光発電設備', '10kW以上', 'うち2,000kW以上', currentMonthRow[7], unit, tab, title, cd(7, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '風力発電設備', '20kW未満', '', currentMonthRow[8], unit, tab, title, cd(8, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '風力発電設備', '20kW以上', '', currentMonthRow[9], unit, tab, title, cd(9, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '風力発電設備', '20kW以上', 'うち洋上風力', currentMonthRow[10], unit, tab, title, cd(10, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW未満', '', currentMonthRow[11], unit, tab, title, cd(11, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW未満', 'うち特定水力', currentMonthRow[12], unit, tab, title, cd(12, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW以上 1,000kW未満', '', currentMonthRow[13], unit, tab, title, cd(13, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '200kW以上 1,000kW未満', 'うち特定水力', currentMonthRow[14], unit, tab, title, cd(14, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '1,000kW以上 30,000kW未満', '', currentMonthRow[15], unit, tab, title, cd(15, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '水力発電設備', '1,000kW以上 30,000kW未満', 'うち特定水力', currentMonthRow[16], unit, tab, title, cd(16, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '地熱発電設備', '15,000kW未満', '', currentMonthRow[17], unit, tab, title, cd(17, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], '地熱発電設備', '15,000kW以上', '', currentMonthRow[18], unit, tab, title, cd(18, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮なし）', 'メタン発酵ガス', '', currentMonthRow[19], unit, tab, title, cd(19, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮なし）', '未利用木質', '', currentMonthRow[20], unit, tab, title, cd(20, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮なし）', '一般木質・農作物残さ', '', currentMonthRow[21], unit, tab, title, cd(21, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮なし）', '建設廃材', '', currentMonthRow[22], unit, tab, title, cd(22, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮なし）', '一般廃棄物・木質以外', '', currentMonthRow[23], unit, tab, title, cd(23, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮あり）', 'メタン発酵ガス', '', currentMonthRow[24], unit, tab, title, cd(24, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮あり）', '未利用木質', '', currentMonthRow[25], unit, tab, title, cd(25, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮あり）', '一般木質・農作物残さ', '', currentMonthRow[26], unit, tab, title, cd(26, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮あり）', '建設廃材', '', currentMonthRow[27], unit, tab, title, cd(27, currentMonthRow, previousMonthRow)))
        rowList.append(Row(year, month, currentMonthRow[0], 'バイオマス発電設備（バイオマス比率考慮あり）', '一般廃棄物・木質以外', '', currentMonthRow[28], unit, tab, title, cd(28, currentMonthRow, previousMonthRow)))

        rowList.append(Row(year, month, currentMonthRow[0], '合計（バイオマス発電設備については、バイオマス比率を考慮したものを合計）', '', '', currentMonthRow[29], unit, tab, title, cd(29, currentMonthRow, previousMonthRow)))
    return rowList
