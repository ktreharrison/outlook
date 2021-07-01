import os
import win32api, win32com.client
from datetime import datetime, timedelta

outlook = win32com.client.Dispatch("Outlook.Application") # initiate outlook client
mapi = outlook.GetNamespace("MAPI")

inbox = mapi.GetDefaultFolder(6)
messages = inbox.Items 

def get_outlook_mailbox():
    return [account.DeliveryStore.DisplayName for account in mapi.Accounts]


def send_outlook_email(to_add, subject, body, attachment=None):
    mail = outlook.CreateItem(0)
    mail.To = to_add 
    mail.Subject = subject 
    mail.HTMLBody = '<h3>This is HTML Body</h3>'
    mail.HTMLBody = body
    mail.Attachments.Add(attachment)
    mail.Send()
    print('email sent')


def get_all_folders(email=None):
    # use email address to access folders:
    if email:
            for idx, folder in enumerate(mapi.Folders(email).Folders):
                print(idx+1, folder)
    else:
        box = int(input("Please select a folder type: \n\n [4]- for folders || [5]- for all folders || [3]-for deleted folder || [6]-standard folders:  "))
        for idx, folder in enumerate(mapi.Folders(box).Folders):
            print(idx+1, folder)



