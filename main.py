from openpyxl import Workbook
from openpyxl import load_workbook

workbook = load_workbook(filename = "Physics-Test.xlsx")
sheet = workbook.active
wrongQ = sheet["D2"].value
wrongQ = wrongQ.replace(" ", "")
wrongQ = wrongQ.split(",")
messages = ""
for count in range(1):
    num = count+2
    for i in wrongQ:
        mnum = str(int(i)+1)
        messages = messages + (sheet["E" +mnum].value) + "  \n"
    sheet["F" + str(num)] = messages
workbook.save(filename= "Physics-Test.xlsx")


