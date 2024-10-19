from openpyxl import Workbook
from openpyxl import load_workbook

workbook = load_workbook(filename = "Physics-Test.xlsx")
sheet = workbook.active

for count in range(1):
    num = count+2
    wrongQ = sheet["D"+ str(num)].value
    wrongQ = wrongQ.replace(" ", "")
    wrongQ = wrongQ.split(",")
    messages = ""


    for i in wrongQ:
        mnum = str(int(i)+1)
        messages = messages + (sheet["E" +mnum].value) + "  \n"
    sheet["F" + str(num)] = messages
workbook.save(filename= "Physics-Test.xlsx")


