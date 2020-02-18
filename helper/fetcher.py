import win32com.client
import os

"""
Goes through outlook and scans all unread mails in the inbox folder.
If it sports the given keywords, it searches for two attachments
in the mail and saves them as text files in the working directory.
"""
def fetch_mailattachments():
    
    outlook = win32com.client.Dispatch("Outlook.Application").getNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    keywords = 'SI Daily Traffic'
    mails = inbox.Items.Restrict("[Unread] = true")
    
    sales_filename = ''
    report_dump_filename = ''

    for mail in mails:
        if keywords in mail.subject:
            try:
                attachments = mail.attachments
                
                sales_filename = str(attachments[0].FileName)
                report_dump_filename = str(attachments[1].FileName[:-3]+'txt')

                attachments[0].SaveAsFile(os.getcwd()+ '\\' + 'input_files\\' + str(attachments[0].FileName))
                attachments[1].SaveAsFile(os.getcwd()+ '\\' + 'input_files\\' + str(attachments[1].FileName[:-3]+'txt'))
                
            except:
                print('Cannot save files from email. Check if the email \
                has attachments or if it is the correct email.')
                exit
    
    return sales_filename, report_dump_filename
