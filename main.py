
from openpyxl import load_workbook
import os

def get_user_input():
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
    to_find = input("What would you like to find?\n")
    to_replace = input("What would you like to replace?\n")

    specific_search = input("\nWould you like to scan the whole document or do a specific search?\nType YES for specific search or type NO to search the whole document: ")
    specific_search = specific_search.lower()
    if "yes" in specific_search:
        sheet_number = input("\nPlease enter the sheet number you would like to do the search: ")
        print("In {} we will find {} and replace it with {}. Only sheet number {} will be searched".format(workbook_name, to_find, to_replace,sheet_number))
        return workbook_name, to_find, to_replace, sheet_number
    else:
        print("In {} we will find {} and replace it with {}. The whole document will be searched".format(workbook_name, to_find, to_replace))
        return workbook_name, to_find, to_replace, None

def info_replacer():

    file_name, info_to_find, info_to_replace_with, sheet_number = get_user_input()
    wb = load_workbook(filename = file_name)
    for sheet in wb:
        print(sheet.title)

if __name__ == '__main__':
    info_replacer()



