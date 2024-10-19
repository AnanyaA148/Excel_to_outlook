from openpyxl import load_workbook
import win32com.client

workbook = load_workbook(filename = "Physics-Test.xlsx")
sheet = workbook.active

for count in range(2):
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




outlook = win32com.client.Dispatch("Outlook.Application")
for count2 in range(2):
    num2 = count2 + 2
    new_mail = outlook.CreateItem(0)
    new_mail.To = sheet["C" + str(num2)].value
    new_mail.Subject = "Wrong Answers"
    new_mail.Body = sheet["F" + str(num2)].value
    new_mail.Send()