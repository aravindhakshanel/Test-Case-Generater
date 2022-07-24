from openpyxl import load_workbook
import json

work_book=load_workbook('config.xlsx')
table=work_book["Table Name"]
column=work_book["Column Name"]
show_filter=work_book["Show filter"]
table_json={}
row_count=2
while table['A'+str(row_count)].value!=None:
    table_json.update({table['A'+str(row_count)].value:{
        "column":{},
        "rearrange column":table['C'+str(row_count)].value,
        "font":table['D'+str(row_count)].value,
        "search bar":table['E'+str(row_count)].value,
        "column visibility":table['F'+str(row_count)].value,
        "pagination":table['G'+str(row_count)].value,
        "show count":table['H'+str(row_count)].value,
        "show filter":table['I'+str(row_count)].value,
        "ag grid":table['J'+str(row_count)].value,
        "period filter":table['K'+str(row_count)].value
        }})
    row_count+=1
row_count=2
for table_name in table_json:
    row_count=2
    while column['A'+str(row_count)].value!=None:
        if column['A'+str(row_count)].value==table_name:
            table_json[table_name]['column'].update({column['B'+str(row_count)].value:{
            "column value":column['C'+str(row_count)].value,
            "hyperlinked":column['D'+str(row_count)].value,
            "sort":column['E'+str(row_count)].value,
            "Search":column['F'+str(row_count)].value,
            "ag grid":column['G'+str(row_count)].value,
            "filter":column['H'+str(row_count)].value,
            "frozen column":column['I'+str(row_count)].value,
            "validation count":column['J'+str(row_count)].value
            }})
            value_count=0
            for value_count in range(0,int(table_json[table_name]['column'][column['B'+str(row_count)].value]["validation count"])):
                table_json[table_name]['column'][column['B'+str(row_count)].value].update({"validation"+str(value_count+1):column[chr(75+value_count)+str(row_count)].value})
                value_count+=1   
        row_count+=1

'''with open('table.json','w') as f:
    json.dump(table_json,f,indent=2)'''

for table_name in table_json:
    row_count=2
    while show_filter['A'+str(row_count)].value!=None:
        if show_filter['A'+str(row_count)].value==table_name:
            for column_name in table_json[table_name]['column']:
                if show_filter['B'+str(row_count)].value==column_name:
                    if show_filter['C'+str(row_count)].value=='search':
                        table_json[table_name].update({"show filter value":{column_name:{"filter type":show_filter['C'+str(row_count)].value,
                        "search format":show_filter['D'+str(row_count)].value,
                        "tab format":show_filter['E'+str(row_count)].value
                        }}})
                    '''else:
                        table_json[table_name].update({"show filter value":{column_name:{"filter type":show_filter['C'+str(row_count)].value,
                        "":show_filter['D'+str(row_count)].value,
                        "tab format":show_filter['E'+str(row_count)].value
                        }}})'''
        row_count+=1

