import win32com.client as win32
import os

def send_report(filename):

    email_list = ''
    with open('email_list.txt') as file:
        for line in file:
            email_list = line 
            break
    
    # -- 

    duration = filename[33:-6]
    dur_arr = duration.split('_')
    month = dur_arr[0]
    year = dur_arr[2]
    
    from_to = dur_arr[1]
    yesterday = from_to.split('-')[1]
    if yesterday == '1' or yesterday == '21':
        from_to = from_to + 'st'
    elif yesterday == '2' or yesterday == '22':
        from_to = from_to + 'nd'
    elif yesterday == '3' or yesterday == '23':
        from_to = from_to + 'rd'
    else:
        from_to = from_to + 'th'

    
    dur_for_email = from_to + ' ' + month + ' ' + year

    # ----

    outlook = win32.Dispatch("Outlook.Application")
    
    mail = outlook.CreateItem(0)
    mail.Subject = 'SI Daily Traffic and Recharge Report ' + dur_for_email
    mail.To = email_list
    mail.HTMLBody = '<p>Automated report generator trial.</p>'
    mail.Body = 'Hi all,\nReport from ' + dur_for_email + ' attached.\n\nLidguard Belo\nICT Department\nMobile: +6778444606\nEmail: lidguard.belo@bmobile.com.sb'

    parent_dir = '\\output\\'
    attachment = parent_dir + filename
    attachment_full_path = os.getcwd() + attachment
    mail.Attachments.Add(attachment_full_path)

    mail.Send()
    # mail.display()

