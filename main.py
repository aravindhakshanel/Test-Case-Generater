import json
import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook




if __name__=="__main__":
    json_str=json.load(open('testrun.json'))
    for senario in json_str:
        for submodule in json_str[senario]:
            for funct in json_str[senario][submodule]:
                for field in json_str[senario][submodule][funct]:
                    for value in json_str[senario][submodule][funct][field]:
                        if value == 'allow':
                            for allow_type in json_str[senario][submodule][funct][field][value]:
                                if allow_type=='module':
                                    print("module testcase")
                                elif allow_type=='checkbox':
                                    print("checkbox")
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