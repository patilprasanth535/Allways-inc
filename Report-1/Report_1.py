import pandas as pd
from xlrd import open_workbook
from xlwt import Workbook, XFStyle
from Styles.Report_styles import ReportStyles
from pandas import ExcelWriter

def main():
    # Here we are reading the raw data csv file
    re_style = ReportStyles()
    csv_file = pd.read_csv('google.csv')
    # Declaring list variables
    month = []
    year = []
    Stock = []
    # assigning values to variable
    stock = csv_file['Stock']
    # looping to extract months and year from date
    for index in csv_file.index:
        #assiging all the date to a variable
        dates = pd.DatetimeIndex(csv_file['date'])
        #appending years to a list
        year.append(dates[index].year)
        #appending months to a list
        month.append(dates[index].month)
        #appending stock to a list
        Stock.append(stock[index])
    #creating a data dictonary
    data = {'Stock': Stock, 'Month': month, 'Year': year}
    #creating a pandas dataframe from dict
    frame = pd.DataFrame(data)
    #Merging two data frames into one common data frame
    csv_file = pd.merge(csv_file, frame, on='Stock')
    #pivoting the data
    piv = csv_file.pivot_table(
        ['Open', 'High'], rows='Month', cols='Year',
        margins=True, aggfunc='count')
    #writing pivot table to an excel
    piv.to_excel('temp.xls')
    book = open_workbook('temp.xls')
    #reading the first sheet from excel
    sheet0 = book.sheet_by_index(0)
    col_cnt = sheet0.ncols
    row_cnt = sheet0.nrows
    pd1 = pd.read_excel(io='temp.xls', sheetname='Sheet1')
    pd2 = pd.read_excel(io='temp.xls', sheetname='Sheet1')
    writer = ExcelWriter('temp1.xls')
    pd1.to_excel(writer,'Sheet1',startcol=0, startrow =2)
    pd2.to_excel(writer,'Sheet1',startcol=(col_cnt+2),startrow =2)
    writer.save()

    book = open_workbook('temp1.xls')
    #reading the first sheet from excel
    sheet0 = book.sheet_by_index(0)
    col_cnt1 = sheet0.ncols
    row_cnt1 = sheet0.nrows

    currency = XFStyle()
    currency.borders = re_style.borders_light()
    currency.alignment = re_style.align_hor_right()
    currency.num_format_str = "[$$-409]#,##0.00;-[$$-409]#,##0.00"

    headings = XFStyle()
    headings.borders = re_style.borders_light()
    headings.alignment = re_style.align_hor_center()
    headings.font = re_style.text_bold()

    no_borders = XFStyle()
    no_borders.borders = re_style.no_borders()


    wb = Workbook()
    ws = wb.add_sheet('Sample_Report', cell_overwrite_ok=True)
    for row in range(row_cnt1):
        for col in range(col_cnt1):
            val = sheet0.cell_value(row, col)
            if row < 2:
                ws.row(row).write(col, val, no_borders)
            elif col == (col_cnt+2):
                ws.row(row).write(col, val, headings)
            # elif col > col_cnt and col < (col_cnt+3):
            #     ws.row(row).write(col, val, no_borders)
            elif row > 4 and col > 0:
                ws.row(row).write(col, val, currency)
            elif row > 4 and col > (col_cnt + 3):
                ws.row(row).write(col, val, currency)
            else:
                ws.row(row).write(col, val, headings)
            wb.save('Report-1.xls')

if __name__ == '__main__':
    main()
