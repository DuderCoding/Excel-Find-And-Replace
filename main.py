
from openpyxl import load_workbook


if __name__ == '__main__':

    wb = load_workbook(filename = "TestExcel.xlsx")

    for sheet in wb:
        print(sheet.title)


