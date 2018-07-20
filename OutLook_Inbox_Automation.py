__project__ = 'OSADOL.'
__author__ = 'Jeys_Ozzius.'
__descript__ = "Copyright (c) 2017 Anirban Ghosh(Anirban83314) & Debashis Biswas(deb991)."
__URL__ = 'https://github.com/Anirban83314/OutLook-Automation-All-scopes-with-APP-utility.'
__NB__ = 'For more information, please see github page & all commit details.'
__CipherSig__ = 'This project & all associate files are encrypted under PBEncryption cryptography. Another details will be available at the end of this pgogram.'

import os
import sys
import win32com.client
import datetime as dt

outlook = win32com.client.Dispatch("Outlook.application")
mapi = outlook.GetNamespace("MAPI")

date_time = dt.datetime.now()
lastQrtDateTime = dt.datetime.now() - dt.timedelta(minutes = 30)

Monitoring_Mails = mapi.Folders['Debashis.Biswas@xyz.com'].Folders['Inbox'].Folders['xyz Monitoring service']
File_Mover_Alerts = mapi.Folders['Debashis.Biswas@xyz.com'].Folders['Inbox'].Folders['File_Mover_Alert_Mail_Folder']

#-----------------Monitoring Items--------------#
MonitoringMails = Monitoring_Mails.Items
MonitoringMails.Sort("[ReceivedTime]", True)
MonitoringMail = MonitoringMails.GetFirst()
MonitoringMail_Count = MonitoringMails.count
MonitoringMail_subject = MonitoringMail.Subject
MonitoringMail_body = MonitoringMail.body
lastQtrMessages = MonitoringMails.Restrict("[ReceivedTime] >= '" +lastQrtDateTime.strftime('%m/%d/%Y %H:%M %p')+"'")

print('/n')
print(MonitoringMail)
print(MonitoringMail_subject)
print(MonitoringMail_body)

print('>>>>>>>>>Monitoring Mails Alert<<<<<<<<')
print('/n')


for MonitoringMail in lastQtrMessages:
    print(MonitoringMail_subject)
    print(MonitoringMail.ReceivedTime)
    MonitoringMail = lastQtrMessages.GetFirst()
    MonitoringMail_iter = lastQtrMessages.GetNext()
    print('/n')

while MonitoringMail:

    for MonitoringMail in lastQtrMessages:
        if MonitoringMail.Unread == True:

            print('Initiating Main process -- flag 1.1')
            if MonitoringMail_subject == 'UnTrusted Message :XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX':
                print('Initiaitng Shell drive partner auto-acknowledgement mail. ')
                exec(open("D:\\OutLook Mail Automation\\xxxxxxxxxxxxxxMailACK.py").read())
                MonitoringMail.Unread == False
                pass


            elif MonitoringMail_subject == 'Something else will be there !!!!':
                print('Bang on again!!!')
                pass

            else:
                print('Nothing important in out Monitoring DL ~!!!!!!!!~~~')



            break
        break
    break

print('EOF')
