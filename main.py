import jpype
import xlsxwriter

from xlsx2html import xlsx2html
import pdfkit


from bs4 import BeautifulSoup


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
    #workbook.save('Expenses01.pdf', SaveFormat.PDF)

    workbook.close()



def trasnf_report_html():
    xlsx2html('./Expenses01.xlsx', './Expenses01.html')
    with open('./Expenses01.html') as f:
        pdfkit.from_file(f, './Expenses01.pdf')

def extrac_table():
    with open('./Expenses01.html') as response:
        content = response.read()
    bs = BeautifulSoup(content, 'html.parser')
    tables = bs.find_all('table')
    return tables[0]


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_excel()
    trasnf_report_html()
    table = extrac_table()


