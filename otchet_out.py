import xlsxwriter
import random
A0 = "Верхняя граница"
A1=10
A2=22
A3=24
A4=24
A5=23
A6=25
A7=23
A8=20
A9=21
A10=27
A11=29
B0 = "Число реализаций"
B1=1
B2=2
B3=2
B4=2
B5=2
B6=2
B7=2
B8=2
B9=2
B10=2
B11=2
# Example data
# Try to do as much processing outside of initializing the workbook
# Everything beetween Workbook() and close() gets trapped in an exception
# random_data = [random.random() for _ in range(10)]
data = [A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11]
data1 = [B1, B2, B3, B4, B5, B6, B7, B8, B9, B10, B11]
# Data location inside excel
data_start_loc = [1, 0] # xlsxwriter rquires list, no tuple
data_end_loc = [data_start_loc[0] + len(data), 0]
data_start_loc1 = [1, 1] # xlsxwriter rquires list, no tuple
data_end_loc1 = [data_start_loc[0] + len(data), 1]

workbook = xlsxwriter.Workbook('file1.xlsx')

# Charts are independent of worksheets
chart = workbook.add_chart({'type': 'line'})
chart.set_y_axis({'name': 'Random jiggly bit values'})
chart.set_x_axis({'name': 'Sequential order'})
chart.set_title({'name': 'Insecure randomly jiggly bits'})

worksheet = workbook.add_worksheet("Задача1")

cell_format = workbook.add_format()
cell_format.set_text_wrap()
worksheet.write(0, 0, "Верхняя граница", cell_format)
worksheet.write(0, 1, "Число реализаций", cell_format)
# A chart requires data to reference data inside excel
worksheet.write_column(*data_start_loc, data=data)     # формирование первого столбца
worksheet.write_column(*data_start_loc1, data=data1)    # формирование второго столбца
# The chart needs to explicitly reference data
chart.add_series({
    'values': [worksheet.name] + data_start_loc + data_end_loc,
    'name': "ЗАДАЧА1",
})
worksheet.insert_chart('C1', chart)


chart = workbook.add_chart({'type': 'line'})
chart.set_y_axis({'name': 'Random jiggly bit values'})
chart.set_x_axis({'name': 'Sequential order'})
chart.set_title({'name': 'Insecure randomly jiggly bits'})
worksheet = workbook.add_worksheet("Задача2")

# A chart requires data to reference data inside excel
worksheet.write_column(*data_start_loc, data=data)     # формирование первого столбца
worksheet.write_column(*data_start_loc1, data=data1)    # формирование второго столбца
# The chart needs to explicitly reference data
chart.add_series({
    'values': [worksheet.name] + data_start_loc + data_end_loc,
    'name': "Random data",
})
worksheet.insert_chart('C1', chart)


chart = workbook.add_chart({'type': 'line'})
chart.set_y_axis({'name': 'Random jiggly bit values'})
chart.set_x_axis({'name': 'Sequential order'})
chart.set_title({'name': 'Insecure randomly jiggly bits'})
worksheet = workbook.add_worksheet("Задача3")

# A chart requires data to reference data inside excel
worksheet.write_column(*data_start_loc, data=data)     # формирование первого столбца
worksheet.write_column(*data_start_loc1, data=data1)    # формирование второго столбца
# The chart needs to explicitly reference data
chart.add_series({
    'values': [worksheet.name] + data_start_loc + data_end_loc,
    'name': "Random data",
})
worksheet.insert_chart('C1', chart)



workbook.close()  # Write to file


