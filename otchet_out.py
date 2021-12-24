from typing import List, Any

import xlsxwriter
import datetime
import random

# A0 = "Верхняя граница"
A1 = 10
A2 = 22
A3 = 24
A4 = 24
A5 = 23
A6 = 25
A7 = 23
A8 = 20
A9 = 21
A10 = 27
A11 = 29
# B0 = "Число реализаций"
B1 = 1
B2 = 2
B3 = 3
B4 = 4
B5 = 5
B6 = 6
B7 = 7
B8 = 8
B9 = 9
B10 = 10
B11 = 11
a = [A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11]
b = [B1, B2, B3, B4, B5, B6, B7, B8, B9, B10, B11]
c = [A1+1, A2+1, A3+1, A4+1, A5+1, A6+1, A7+1, A8+1, A9+1, A10+1, A11+1]
d = [B1-1, B2-1, B3-1, B4-1, B5-1, B6-1, B7-1, B8-1, B9-1, B10-1, B11-1]
e = [A1*2, A2*2, A3*2, A4*2, A5*2, A6*2, A7*2, A8*2, A9*2, A10*2, A11*2]
f = [B1, B2, B3, B4, B5, B6, B7, B8, B9, B10, B11]
g = [A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11]
h = [B1, B2, B3, B4, B5, B6, B7, B8, B9, B10, B11]
i = [A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11]
j = [B1, B2, B3, B4, B5, B6, B7, B8, B9, B10, B11]
k = [A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11]
l = [B1, B2, B3, B4, B5, B6, B7, B8, B9, B10, B11]
# a = [A1, A2, A3, A4]
# b = [B1, B2, B3, B4]

def create_report(a1, b1, c1, d1, a2, b2, c2, d2, a3, b3, c3, d3):
    now = datetime.datetime.now()
    dat = now.strftime("%d-%m-%Y %H")
    # Example data
    # Try to do as much processing outside of initializing the workbook
    # Everything beetween Workbook() and close() gets trapped in an exception
    data = a1
    data1 = b1
    data2 = c1
    data3 = d1
    data4 = a2
    data5 = b2
    data6 = c2
    data7 = d2
    data8 = a3
    data9 = b3
    data10 = c3
    data11 = d3
    # Data location inside excel
    data_start_loc = [1, 0]  # xlsxwriter rquires list, no tuple
    data_end_loc = [data_start_loc[0] + len(data), 0]
    data_start_loc1 = [1, 1]  # xlsxwriter rquires list, no tuple
    data_end_loc1 = [data_start_loc1[0] + len(data1), 1]
    data_start_loc2 = [1, 10]  # xlsxwriter rquires list, no tuple
    data_end_loc2 = [data_start_loc2[0] + len(data2), 0]
    data_start_loc3 = [1, 11]  # xlsxwriter rquires list, no tuple
    data_end_loc3 = [data_start_loc3[0] + len(data3), 1]
    data_start_loc4 = [1, 0]  # xlsxwriter rquires list, no tuple
    data_end_loc4 = [data_start_loc4[0] + len(data4), 0]
    data_start_loc5 = [1, 1]  # xlsxwriter rquires list, no tuple
    data_end_loc5 = [data_start_loc5[0] + len(data5), 1]
    data_start_loc6 = [1, 10]  # xlsxwriter rquires list, no tuple
    data_end_loc6 = [data_start_loc6[0] + len(data6), 0]
    data_start_loc7 = [1, 11]  # xlsxwriter rquires list, no tuple
    data_end_loc7 = [data_start_loc7[0] + len(data7), 1]
    data_start_loc8 = [1, 0]  # xlsxwriter rquires list, no tuple
    data_end_loc8 = [data_start_loc8[0] + len(data8), 0]
    data_start_loc9 = [1, 1]  # xlsxwriter rquires list, no tuple
    data_end_loc9 = [data_start_loc9[0] + len(data9), 1]
    data_start_loc10 = [1, 10]  # xlsxwriter rquires list, no tuple
    data_end_loc10 = [data_start_loc10[0] + len(data10), 0]
    data_start_loc11 = [1, 11]  # xlsxwriter rquires list, no tuple
    data_end_loc11 = [data_start_loc11[0] + len(data11), 1]

    workbook = xlsxwriter.Workbook(f'report_{dat}.xlsx')

    # Charts are independent of worksheets
    chart = workbook.add_chart({'type': 'line'})
    chart.set_y_axis({'name': 'Верхняя граница'})
    chart.set_x_axis({'name': 'Число реализаций'})
    chart.set_title({'name': 'Результат испытаний'})

    worksheet = workbook.add_worksheet("Задача1")

    cell_format = workbook.add_format()
    cell_format.set_text_wrap()
    cell_format.set_align('center')
    cell_format.set_align('vcenter')
    worksheet.write(0, 0, "Верхняя граница", cell_format)
    worksheet.write(0, 1, "Число реализаций", cell_format)
    # worksheet.write(14, 0, dat, cell_format)
    # A chart requires data to reference data inside excel
    worksheet.write_column(*data_start_loc, data=data)  # формирование первого столбца
    worksheet.write_column(*data_start_loc1, data=data1)  # формирование второго столбца
    # The chart needs to explicitly reference data
    chart.add_series({
        'values': [worksheet.name] + data_start_loc + data_end_loc,
        'name': "data",
    })
    worksheet.insert_chart('C1', chart)

    # Charts are independent of worksheets
    chart = workbook.add_chart({'type': 'line'})
    chart.set_y_axis({'name': 'Y'})
    chart.set_x_axis({'name': 'X'})
    chart.set_title({'name': 'График 2'})


    cell_format = workbook.add_format()
    cell_format.set_text_wrap()
    cell_format.set_align('center')
    cell_format.set_align('vcenter')
    worksheet.write(0, 10, "m", cell_format)
    worksheet.write(0, 11, "E", cell_format)
    # worksheet.write(14, 0, dat, cell_format)
    # A chart requires data to reference data inside excel
    worksheet.write_column(*data_start_loc2, data=data2)  # формирование первого столбца
    worksheet.write_column(*data_start_loc3, data=data3)  # формирование второго столбца
    # The chart needs to explicitly reference data
    chart.add_series({
        # 'values': [worksheet.name] + data_start_loc2 + data_end_loc2,
        'values': f'=Задача1!$K$1:$K${len(data2)}',
        'name': "data1",
    })
    worksheet.insert_chart('M1', chart)

    chart = workbook.add_chart({'type': 'line'})
    chart.set_y_axis({'name': 'Верхняя граница'})
    chart.set_x_axis({'name': 'Число реализаций'})
    chart.set_title({'name': 'Результат испытаний'})

    worksheet = workbook.add_worksheet("Задача2")
    cell_format = workbook.add_format()
    cell_format.set_text_wrap()
    cell_format.set_align('center')
    cell_format.set_align('vcenter')
    worksheet.write(0, 0, "Верхняя граница", cell_format)
    worksheet.write(0, 1, "Число реализаций", cell_format)

    # A chart requires data to reference data inside excel
    worksheet.write_column(*data_start_loc4, data=data4)  # формирование первого столбца
    worksheet.write_column(*data_start_loc5, data=data5)  # формирование второго столбца
    # The chart needs to explicitly reference data
    chart.add_series({
        'values': [worksheet.name] + data_start_loc + data_end_loc,
        'name': "data",
    })
    worksheet.insert_chart('C1', chart)
    chart = workbook.add_chart({'type': 'line'})
    chart.set_y_axis({'name': 'Y'})
    chart.set_x_axis({'name': 'X'})
    chart.set_title({'name': 'График 2'})

    cell_format = workbook.add_format()
    cell_format.set_text_wrap()
    cell_format.set_align('center')
    cell_format.set_align('vcenter')
    worksheet.write(0, 10, "m", cell_format)
    worksheet.write(0, 11, "E", cell_format)
    # worksheet.write(14, 0, dat, cell_format)
    # A chart requires data to reference data inside excel
    worksheet.write_column(*data_start_loc6, data=data6)  # формирование первого столбца
    worksheet.write_column(*data_start_loc7, data=data7)  # формирование второго столбца
    # The chart needs to explicitly reference data
    chart.add_series({
        # 'values': [worksheet.name] + data_start_loc2 + data_end_loc2,
        'values': f'=Задача2!$K$1:$K${len(data2)}',
        'name': "data1",
    })
    worksheet.insert_chart('M1', chart)

    chart = workbook.add_chart({'type': 'line'})
    chart.set_y_axis({'name': 'Верхняя граница'})
    chart.set_x_axis({'name': 'Число реализаций'})
    chart.set_title({'name': 'Результат испытаний'})

    worksheet = workbook.add_worksheet("Задача3")
    cell_format = workbook.add_format()
    cell_format.set_text_wrap()
    cell_format.set_align('center')
    cell_format.set_align('vcenter')
    worksheet.write(0, 0, "Верхняя граница", cell_format)
    worksheet.write(0, 1, "Число реализаций", cell_format)

    # A chart requires data to reference data inside excel
    worksheet.write_column(*data_start_loc8, data=data8)  # формирование первого столбца
    worksheet.write_column(*data_start_loc9, data=data9)  # формирование второго столбца
    # The chart needs to explicitly reference data
    chart.add_series({
        'values': [worksheet.name] + data_start_loc + data_end_loc,
        'name': "data",
    })
    worksheet.insert_chart('C1', chart)
    chart = workbook.add_chart({'type': 'line'})
    chart.set_y_axis({'name': 'Y'})
    chart.set_x_axis({'name': 'X'})
    chart.set_title({'name': 'График 2'})

    cell_format = workbook.add_format()
    cell_format.set_text_wrap()
    cell_format.set_align('center')
    cell_format.set_align('vcenter')
    worksheet.write(0, 10, "m", cell_format)
    worksheet.write(0, 11, "E", cell_format)
    # worksheet.write(14, 0, dat, cell_format)
    # A chart requires data to reference data inside excel
    worksheet.write_column(*data_start_loc10, data=data10)  # формирование первого столбца
    worksheet.write_column(*data_start_loc11, data=data11)  # формирование второго столбца
    # The chart needs to explicitly reference data
    chart.add_series({
        # 'values': [worksheet.name] + data_start_loc2 + data_end_loc2,
        'values': f'=Задача3!$K$1:$K${len(data10)}',
        'name': "data1",
    })
    worksheet.insert_chart('M1', chart)

    workbook.close()  # Write to file


create_report(a, b, c, d, e, f, g, h, i, j, k, l)
