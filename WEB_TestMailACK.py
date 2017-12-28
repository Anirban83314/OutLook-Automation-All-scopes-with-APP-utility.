#!/usr/bin/env python
import win32com.client as win32

from MasterAlertMonitor import _Mail_body
from MasterAlertMonitor import _Mail_subject
import web_Test

# , CTM_Mail_body

outlook = win32.Dispatch ( 'outlook.application' )
mail = outlook.CreateItem ( 0 )
mail.To = 'receiver'
mail.Cc = 'receiver'
mail.Bcc = 'receiver'
mail.Subject = 'Acknowledgement || Re: ' + _Mail_subject

attachment = ('C:\\Users\\Debashis.Biswas\\Desktop\\WebModules\\' + web_AliveTest.newest)
mail.Attachments.Add(attachment)

(mail.body) = 'Hence Web UI has been checked as per the requirement & acknowledged. ' + _Mail_body

mail.send
