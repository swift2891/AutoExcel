from sys import argv
import openpyxl
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles import colors
from openpyxl.cell import Cell
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.chart import (
    ScatterChart,
    Reference,
    Series,
)

class AutoExcel:
    # Input Excel
    print("Loading Input file: " + argv[1])
    wb = openpyxl.load_workbook(argv[1])
    # Output Excel
    wb_O = openpyxl.Workbook()
    # Sheet1
    sh1 = wb.get_sheet_by_name('Sheet1')
    maxRows = sh1.max_row
    sh1_O = wb_O.get_active_sheet()

    groupLength = []
    columnStart = 'A'
    index = 1
    prevNegative = 'N'

    @staticmethod
    def negativeProc(current,i):
        AutoExcel.prevNegative = 'Y'

        potentialCoordinate = AutoExcel.columnStart + str(AutoExcel.index)
        currentCoord = get_column_letter(int(column_index_from_string(AutoExcel.columnStart)) + 1) + str(AutoExcel.index)
        timeCoord = get_column_letter(int(column_index_from_string(AutoExcel.columnStart)) + 2) + str(AutoExcel.index)
        capacityCoord = get_column_letter(int(column_index_from_string(AutoExcel.columnStart)) + 3) + str(AutoExcel.index)

        print("the coordinates are: " + potentialCoordinate + " " + currentCoord + " " + capacityCoord)
        potential = AutoExcel.sh1.cell(row=i, column=1)
        time = AutoExcel.sh1.cell(row=i, column=3)
        capacity = AutoExcel.sh1.cell(row=i, column=10)

        AutoExcel.sh1_O[potentialCoordinate] = potential.value
        AutoExcel.sh1_O[currentCoord] = current.value
        AutoExcel.sh1_O[timeCoord] = time.value
        AutoExcel.sh1_O[capacityCoord] = capacity.value
        AutoExcel.index += 1
        print("Row#: " + str(i) + "  Saving " + str(current.value) + " in output..")
        if i == AutoExcel.maxRows:
            AutoExcel.groupLength.append(AutoExcel.index - 1)

    @staticmethod
    def addChart(groupLength):
        print('Printing Chart...')
        chart = ScatterChart()
        chart.title = "Capacity Vs Voltage"
        chart.style = 13
        chart.x_axis.title = 'Capacity'
        chart.y_axis.title = 'Voltage'
        le = (len(groupLength) * 4) + len(groupLength)
        print("The sizes of each grp" + str(groupLength[0]))
        k = 0
        m = 1
        print("list siz = " + str(len(groupLength)))
        print("le = " + str(le))
        for j in range(4, le, 5):
            print("j = " + str(j))
            xvalues = Reference(AutoExcel.sh1_O, min_col=j, min_row=1, max_row=groupLength[k])
            values = Reference(AutoExcel.sh1_O, min_col=m, min_row=1, max_row=groupLength[k])
            series = Series(values, xvalues, title_from_data=True)
            chart.series.append(series)
            k += 1
            m += 5

        AutoExcel.sh1_O.add_chart(chart, 'L3')
        AutoExcel.wb_O.save('output.xlsx')

    @staticmethod
    def mainApp():
        redFill = PatternFill(start_color='FFFF0000',
                              end_color='FFFF0000',
                              fill_type='solid')
        for i in range(2, AutoExcel.maxRows+1):
            current = AutoExcel.sh1.cell(row=i,column=2)
            # None eliminator
            if current.value == None:
                current.value = 0
            if current.value<0:
                AutoExcel.negativeProc(current,i)
            elif AutoExcel.prevNegative == 'Y':
                AutoExcel.prevNegative = 'N'
                potentialCoordinate = AutoExcel.columnStart + str(AutoExcel.index)
                currentCoord = get_column_letter(int(column_index_from_string(AutoExcel.columnStart)) + 1) + str(AutoExcel.index)
                timeCoord = get_column_letter(int(column_index_from_string(AutoExcel.columnStart)) + 2) + str(AutoExcel.index)
                capacityCoord = get_column_letter(int(column_index_from_string(AutoExcel.columnStart)) + 3) + str(AutoExcel.index)
                AutoExcel.sh1_O[currentCoord].fill = redFill
                AutoExcel.sh1_O[potentialCoordinate].fill = redFill
                AutoExcel.sh1_O[timeCoord].fill = redFill
                AutoExcel.sh1_O[capacityCoord].fill = redFill
                AutoExcel.groupLength.append(AutoExcel.index-1)
                AutoExcel.index = 1
                AutoExcel.columnStart = get_column_letter(int(column_index_from_string(AutoExcel.columnStart))+5)
            else:
                continue
        print("Saving values..")
        AutoExcel.wb_O.save('output.xlsx')

        AutoExcel.addChart(AutoExcel.groupLength)

        print("xxxxx------ Program Ended ------xxx")
        print("Output File Created: output.xlsx")

AutoExcel.mainApp()