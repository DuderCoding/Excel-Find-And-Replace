
from openpyxl import load_workbook
import os

def info_replacer():

    file_exists = False
    while not file_exists:
        workbook_name = input("Please enter Excel file name: ")
        if not ".xlsx" in workbook_name:
            workbook_name = workbook_name + ".xlsx"

        file_exists = os.path.exists(workbook_name)

        if file_exists:
            print("File found, thank you...\n")
        else:
            print("{} file not found in folder\n".format(workbook_name))

    wb = load_workbook(filename = "TestExcel.xlsx")
    for sheet in wb:
        print(sheet.title)

if __name__ == '__main__':
    info_replacer()



