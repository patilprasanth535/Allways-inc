import pandas as pd
from xlrd import open_workbook
from xlwt import Workbook, XFStyle
from Report_styles import ReportStyles

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
    piv.to_excel('C:\Users\prasanth\Desktop\gg.xls')
    #opening the excel file
    book = open_workbook('C:\Users\prasanth\Desktop\gg.xls')
    #reading the first sheet from excel
    sheet0 = book.sheet_by_index(0)

    currency = XFStyle()
    currency.borders = re_style.borders_light()
    currency.alignment = re_style.align_hor_right()
    currency.num_format_str = "[$$-409]#,##0.00;-[$$-409]#,##0.00"

    headings = XFStyle()
    headings.borders = re_style.borders_light()
    headings.alignment = re_style.align_hor_center()
    headings.font = re_style.text_bold()

    col_cnt = sheet0.ncols
    row_cnt = sheet0.nrows
    wb = Workbook()
    ws = wb.add_sheet('Type examples', cell_overwrite_ok=True)
    for row in range(row_cnt):
        for col in range(col_cnt):
            val = sheet0.cell_value(row, col)
            if row > 2 and col > 0:
                ws.row(row).write(col, val, currency)
            else:
                ws.row(row).write(col, val, headings)
            wb.save('Report-1.xls')

if __name__ == '__main__':
    main()
