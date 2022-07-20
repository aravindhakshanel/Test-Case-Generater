import json
from multiprocessing.spawn import get_command_line
import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter




if __name__=="__main__":
    #json file
    json_str=json.load(open('testrun.json'))
    #config file
    work_book=load_workbook('config.xlsx')
    column_name=work_book['Column Name']
    table_name=work_book['Table Name']
    allow =work_book['Allow']
    Show_filter=work_book['Show filter']
    ag_grid=work_book['Ag grid']
    general=work_book['General']
    #Test case sheet
    work_book2=load_workbook('test case.xlsx')
    output_sheet=work_book2['Sheet1']
    row_count=2
    for senario in json_str:
        for submodule in json_str[senario]:
            for funct in json_str[senario][submodule]:
                for field in json_str[senario][submodule][funct]:
                    for value in json_str[senario][submodule][funct][field]:
                        if value == 'allow':
                            for allow_type in json_str[senario][submodule][funct][field][value]:
                                if allow_type=='module':
                                    output_sheet['H'+str(row_count)]=senario
                                    output_sheet['E'+str(row_count)]=submodule
                                    output_sheet['F'+str(row_count)]=funct
                                    output_sheet['G'+str(row_count)]=field
                                    output_sheet['O'+str(row_count)]='Check whether the application allow to click '+json_str[senario][submodule][funct][field][value][allow_type]+' module'
                                    output_sheet['Q'+str(row_count)]='Should display '+json_str[senario][submodule][funct][field][value][allow_type]+' page'
                                    output_sheet['AD'+str(row_count)]='Positive'
                                    row_count=row_count+1
                                elif allow_type=='checkbox':
                                    output_sheet['H'+str(row_count)]=senario
                                    output_sheet['E'+str(row_count)]=submodule
                                    output_sheet['F'+str(row_count)]=funct
                                    output_sheet['G'+str(row_count)]=field
                                    output_sheet['O'+str(row_count)]='Check whether the application allow to check '+json_str[senario][submodule][funct][field][value][allow_type]+' checkbox'
                                    output_sheet['Q'+str(row_count)]='Should display '+json_str[senario][submodule][funct][field][value][allow_type]+' page'
                                    output_sheet['AD'+str(row_count)]='Positive'
                                    row_count=row_count+1
                                    output_sheet['H'+str(row_count)]=senario
                                    output_sheet['E'+str(row_count)]=submodule
                                    output_sheet['F'+str(row_count)]=funct
                                    output_sheet['G'+str(row_count)]=field
                                    output_sheet['O'+str(row_count)]='Check whether the application allow to check '+json_str[senario][submodule][funct][field][value][allow_type]+' checkbox'
                                    output_sheet['Q'+str(row_count)]='Should display '+json_str[senario][submodule][funct][field][value][allow_type]+' page'
                                    output_sheet['AD'+str(row_count)]='Positive'
                                elif allow_type=='icon':
                                    print("edit")
                                elif allow_type=='link':
                                    print("link")
                                elif allow_type=='tab':
                                    print("tab")
                        elif value=='table':
                            print("table")
                        elif value=='general':
                            print('general')
    work_book2.save('test case.xlsx')