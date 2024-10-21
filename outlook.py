# NOT A NECCESARY FILE. Just used to test the library
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")

new_mail = outlook.CreateItem(0)

new_mail.To = "an953660@ucf.edu"

new_mail.Subject = "Wrong Answers"

new_mail.Body = "Emailbody"

new_mail.Send()
