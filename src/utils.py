import xlrd
import xlwt


def get_sheets(file_name):

    return xlrd.open_workbook(file_name).sheets()


def get_columns(sheet, column_headers):
    columns = {}

    for header in column_headers:
        for col in range(sheet.ncols):
            if header in sheet.cell(0, col).value:
                columns[header] = col
    return columns


def save_data(file_name, sheet_names, sheet_data):
    workbook = xlwt.Workbook()

    sheets = []
    for name in sheet_names:
        sheets.append(workbook.add_sheet(name, cell_overwrite_ok=True))

    for num in range(len(sheets)):
        rows = len(sheet_data[num])
        for i in range(rows):
            cols = len(sheet_data[num][i])
            for j in range(cols):
                sheets[num].write(i, j, sheet_data[num][i][j])

    workbook.save(file_name)
