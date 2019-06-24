import openpyxl
import os


prospect_list = []
WRKBOOK = openpyxl.load_workbook('Key Accounts VP Level 2019-06-24.xlsx')
SHEET = WRKBOOK['GetProspect Leads']
final_string = ''


def get_prospect_list():
    for i in range(2, 169):
        prospect_first_name = SHEET.cell(row=i, column=1).value
        prospect_last_name = SHEET.cell(row=i, column=2).value
        prospect_name_formatted = '(' + '"' + prospect_first_name + '"' + ' AND ' + '"' + prospect_last_name + '"' + ")"
        #prospect_name_formatted = '(' + prospect_first_name + ' ' + prospect_last_name + ")"
        #prospect_full_name = str(prospect_first_name) + ' ' + str(prospect_last_name)
        prospect_list.append(prospect_name_formatted)


def search_query_builder():
    search_string = ''
    for i in prospect_list:
        search_string += str(i) + ' OR '
    print(search_string)
    f = open('search-string.txt', "w")
    f.write(search_string)
    f.close()

if __name__ == '__main__':
    get_prospect_list()
    search_query_builder()
