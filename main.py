from openpyxl import Workbook
from openpyxl import load_workbook

workbook = load_workbook(filename = "Physics-Test.xlsx")
sheet = workbook.active
wrongQ = sheet["D2"].value
wrongQ = wrongQ.replace(" ", "")
wrongQ = wrongQ.split(",")
for i in wrongQ:
    num = "E" + str(int(i)+1)
    print(sheet[num].value)


