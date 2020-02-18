import win32com.client as win32
import os

def send_report(filename):

    email_list = ''
    with open('email_list.txt') as file:
        for line in file:
            email_list = line 
            print(email_list)
            break
    
    outlook = win32.Dispatch("Outlook.Application")
    
    mail = outlook.CreateItem(0)
    mail.Subject = 'Test Report Automated Sender'
    mail.To = email_list
    mail.HTMLBody = '<p>Automated report generator trial.</p>'
    attachment = '\\output\\' + filename
    attachment_full_path = os.getcwd() + attachment
    mail.Attachments.Add(attachment_full_path)

    mail.Send()

