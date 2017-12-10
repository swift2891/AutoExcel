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

def mainApp():
    # Input Excel
    print("Loading Input file: "+argv[1])
    wb = openpyxl.load_workbook(argv[1])

    # Output Excel
    wb_O = openpyxl.Workbook()
    sh1_O = wb_O.get_active_sheet()

    # Sheet1
    sh1 = wb.get_sheet_by_name('Sheet1')
    maxRows = sh1.max_row

    redFill = PatternFill(start_color='FFFF0000',
                          end_color='FFFF0000',
                          fill_type='solid')

    groupLength = []
    columnStart = 'A'
    index = 1
    prevNegative = 'N'
    for i in range(2, maxRows+1):
        current = sh1.cell(row=i,column=2)
        # None eliminator
        if current.value == None:
            current.value = 0
        if current.value<0:
            prevNegative = 'Y'

            potentialCoordinate = columnStart + str(index)
            currentCoord = get_column_letter(int(column_index_from_string(columnStart))+1) + str(index)
            capacityCoord = get_column_letter(int(column_index_from_string(columnStart)) + 2) + str(index)

            print("the coordinates are: "+potentialCoordinate +" "+ currentCoord +" "+ capacityCoord)
            # currentCoord = 'B'+str(index)
            potential = sh1.cell(row=i, column=1)
            capacity = sh1.cell(row=i, column=10)

            sh1_O[potentialCoordinate] = potential.value
            sh1_O[currentCoord] = current.value
            sh1_O[capacityCoord]= capacity.value
            index += 1
            print("Row#: "+str(i)+"  Saving "+str(current.value)+" in output..")
            if i==maxRows:
                groupLength.append(index-1)
        elif prevNegative == 'Y':
            prevNegative = 'N'
            potentialCoordinate = columnStart + str(index)
            currentCoord = get_column_letter(int(column_index_from_string(columnStart)) + 1) + str(index)
            capacityCoord = get_column_letter(int(column_index_from_string(columnStart)) + 2) + str(index)
            sh1_O[currentCoord].fill = redFill
            sh1_O[potentialCoordinate].fill = redFill
            sh1_O[capacityCoord].fill = redFill
            groupLength.append(index-1)
            index = 1
            columnStart = get_column_letter(int(column_index_from_string(columnStart))+4)
        else:
            continue
    print("Saving values..")
    wb_O.save('output.xlsx')

    print('Printing Chart...')
    chart = ScatterChart()
    chart.title = "Capacity Vs Voltage"
    chart.style = 13
    chart.x_axis.title = 'Capacity'
    chart.y_axis.title = 'Voltage'
    le = (len(groupLength)*3) + len(groupLength)
    print("The sizes of each grp"+str(groupLength[0]))

    k=0
    m=1
    graphIndex = 3
    print("list siz = " + str(len(groupLength)))
    print("le = " + str(le))
    for j in range(3,le,4):
        print("j = "+str(j))
        xvalues = Reference(sh1_O, min_col=j, min_row=1, max_row=groupLength[k])
        values = Reference(sh1_O, min_col=m, min_row=1, max_row=groupLength[k])
        series = Series(values, xvalues, title_from_data=True)
        chart.series.append(series)
        k += 1
        m+=4

    sh1_O.add_chart(chart, 'L3')
    wb_O.save('output.xlsx')
    print("xxxxx------ Program Ended ------xxx")
    print("Output File Created: output.xlsx")

mainApp()