from openpyxl import load_workbook

work_book=load_workbook('config.xlsx')
allow=work_book["Allow"]
config_json={"button":{},"check box":{},"icon":{},"link":{},"tab":{}}
conf={}
row_count=2
while allow['Q'+str(row_count)].value!=None:
    conf.update({allow['Q'+str(row_count)].value:{"message":allow['R'+str(row_count)].value,"yes":allow['S'+str(row_count)].value,"No":allow['T'+str(row_count)].value}})
    row_count+=1
row_count=2
while allow['A'+str(row_count)].value!=None:
    config_json["button"].update({allow['A'+str(row_count)].value:allow['B'+str(row_count)].value})
    row_count+=1
row_count=2
while allow['D'+str(row_count)].value!=None:
    config_json["check box"].update({allow['D'+str(row_count)].value:{"checked":allow['E'+str(row_count)].value,"unchecked":allow['F'+str(row_count)].value}})
    row_count+=1
row_count=2
while allow['H'+str(row_count)].value!=None:
    config_json["icon"].update({allow['H'+str(row_count)].value:allow['I'+str(row_count)].value})
    row_count+=1
row_count=2
while allow['K'+str(row_count)].value!=None:
    config_json["link"].update({allow['K'+str(row_count)].value:allow['K'+str(row_count)].value})
    row_count+=1
row_count=2
while allow['N'+str(row_count)].value!=None:
    config_json["tab"].update({allow['N'+str(row_count)].value:allow['O'+str(row_count)].value})
    row_count+=1
print(conf)
print(config_json)

