import xlsxwriter
import io
from xlsx2html import xlsx2html
import re


def print_excel():
    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook('Expenses01.xlsx')
    worksheet = workbook.add_worksheet()

    # Some data we want to write to the worksheet.
    expenses = (
        ['Rent', 1000],
        ['Gas', 100],
        ['Food', 300],
        ['Gym', 50],
    )

    # Start from the first cell. Rows and columns are zero indexed.
    row = 0
    col = 0

    # Iterate over the data and write it out row by row.
    for item, cost in (expenses):
        worksheet.write(row, col, item)
        worksheet.write(row, col + 1, cost)
        row += 1

    # Write a total using a formula.
    worksheet.write(row, 0, 'Total')
    worksheet.write(row, 1, '=SUM(B1:B4)')

    workbook.close()


def trasnf_report_html():
    # must be binary mode
    # xlsx_file = open('/home/beto/Documents/Projects/Pruebra-reporte/Expenses01.xlsx', 'rb')
    # out_file = io.StringIO()
    # xlsx2html(xlsx_file, out_file, locale='en')
    # out_file.seek(0)
    # result_html = out_file.read()
    # print(result_html)
    xlsx2html('./Expenses01.xlsx', './Expenses01.html')

def extrac_table():
    test_str = open('./Expenses01.html', 'rb')
    tag = "b"
    print(test_str)

    reg_str = "<" + tag + ">(.*?)</" + tag + ">"
    res = re.findall(reg_str, test_str)

    print("The Strings extracted : " + res)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':

    print_excel()
    trasnf_report_html()
    extrac_table()



# See PyCharm help at https://www.jetbrains.com/help/pycharm/
