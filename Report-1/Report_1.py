import pandas as pd
import numpy as np
from xlrd import open_workbook
from xlwt import Workbook, XFStyle, Borders, Pattern, Font, Style
csv_file = pd.read_csv('C:\Users\prasanth\Desktop\google.csv')
month = []
year = []
Stock = []

stock = csv_file['Stock']
for index in csv_file.index:
    dates = pd.DatetimeIndex(csv_file['date'])
    year.append(dates[index].year)
    month.append(dates[index].month)
    Stock.append(stock[index])
data = {'Stock': Stock, 'Month': month, 'Year': year}
frame = pd.DataFrame(data)
# print frame
csv_file = pd.merge(csv_file, frame, on='Stock')
# print csv_file
piv = csv_file.pivot_table(['Open', 'High'], rows='Month', cols='Year', aggfunc='sum')
print piv
piv.to_excel('C:\Users\prasanth\Desktop\gg.xls')
hgfh
book = open_workbook('C:\Users\prasanth\Desktop\gg.xls')
sheet0 = book.sheet_by_index(0)
borders = Borders()
borders.left = Borders.THIN
borders.right = Borders.THIN
borders.top = Borders.THIN
borders.bottom = Borders.THIN
# pattern = Pattern()
# pattern.pattern = Pattern.SOLID_PATTERN
# pattern.pattern_fore_colour = 0x0A
style = XFStyle()
# style.num_format_str='YYYY-MM-DD'
# style.font = fnt
style.borders = borders
# style.pattern = pattern

col_cnt = sheet0.ncols
row_cnt = sheet0.nrows
wb = Workbook()
ws = wb.add_sheet('Type examples', cell_overwrite_ok=True)
for col in range(col_cnt):
    for row in range(row_cnt):
        val = sheet0.cell_value(row, col)
        ws.row(row).write(col, val, style)
        wb.save('hello.xls')


