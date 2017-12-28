#!/usr/bin/env python
import win32com.client as win32

from FPSAlertMonitor import CTM_Mail_body
from FPSAlertMonitor import CTM_Mail_subject
import PollAliveTest

# , CTM_Mail_body

outlook = win32.Dispatch ( 'outlook.application' )
mail = outlook.CreateItem ( 0 )
mail.To = 'GXSOMWIPRODSMonitoringservice@shell.com'
mail.Cc = 'anirban.a.ghosh@shell.com'
mail.Bcc = 'debashis.biswas@shell.com'
mail.Subject = 'Acknowledgement || Re: ' + CTM_Mail_subject

attachment = ('C:\\Users\\Debashis.Biswas\\Desktop\\WebModules\\' + PollAliveTest.newest)
mail.Attachments.Add(attachment)

(mail.body) = 'Hence Web UI has been checked as per the requirement & acknowledged. ' + CTM_Mail_body

mail.send