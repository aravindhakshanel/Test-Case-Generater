import array
class test_case():
    def __init__(self):
        print('Started')
    def create_test_case(self,i):
        self.test_case_count=i
        self.date="date"
        self.version="version"
        self.brd_id="brd id"
        self.module="module name"
        self.submodule="sub module name"
        self.functionality="functionality"
        self.field="field"
        self.senario="senario"
        self.senario_id="senario id"
        self.manual_id="manual id"
        self.unique="unique id"
        self.description="description"
        self.test_data="test data"
        self.expected_result="expexted result"
        self.type="type"
    def print_test_case(self):
        print(self.test_case_count,self.date,self.version,self.brd_id,self.module)
test=[]
for i in range(0,10):
    test.append(test_case())
    test[i].create_test_case(i)
    test[i].print_test_case()