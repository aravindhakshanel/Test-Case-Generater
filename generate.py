from openpyxl import load_workbook
import json

work_book=load_workbook('config.xlsx')
senario=work_book['Scenario']
i=1
json_gen={}
while senario['A'+str(i)].value!=None:
    json_gen.update({senario['A'+str(i)].value:{}})
    i=i+1
with open('senario.json','w') as f:
    json.dump(json_gen,f,indent=2)

submodule=work_book['Submodule']
i=1
json_gen2={}
while submodule['A'+str(i)].value!=None:
    json_gen2.update({submodule['A'+str(i)].value:{}})
    i=i+1
with open('submodule.json','w') as f:
    json.dump(json_gen2,f,indent=2)

functionality=work_book['Functionality']
i=1
json_gen3={}
while functionality['A'+str(i)].value!=None:
    json_gen3.update({functionality['A'+str(i)].value:{}})
    i=i+1
with open('functionality.json','w') as f:
    json.dump(json_gen3,f,indent=2)