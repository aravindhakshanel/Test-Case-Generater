class excel():
    config_json={"button":{},"check box":{},"icon":{},"link":{},"tab":{}}
    def __init__(self,file_name,json_file):
        from openpyxl import load_workbook
        work_book=load_workbook(file_name)
        self.allow=work_book['Allow']
        self.column=work_book['Column Name']
        self.table=work_book['Table Name']
        self.show_filter=work_book['Show filter']
        self.ag_grid=work_book['Ag grid']
        self.json_valuejson.load(open(json_file))

    def scanfile(self):
        self.button('button_name')

    def button(self,button_name):
        i=2
        while self.sheet['A'+str(i)].value!=None:
            if self.sheet['A'+str(i)].value==button_name:
                return self.sheet['B'+str(i)].value
            i+=1
    def conf(self,popup_name):
        i=2
        while self.sheet['Q'+str(i)].value!=None:
            if self.sheet['Q'+str(i)].value==popup_name:
                return self.sheet['Q'+str(i)].value
            i+=1
    def export(self):
        print("sucess")

p=excel('config.xlsx','testrun.json')
p.scanfile()
p.scanjson()
p.export()
