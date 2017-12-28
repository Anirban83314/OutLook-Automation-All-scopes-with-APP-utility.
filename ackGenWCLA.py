import win32com.client as win32

import FPSAlertMonitor



outlook = win32.Dispatch ( 'outlook.application' )
mail = outlook.CreateItem ( 0 )
mail.To = 'GXSOMWIPRODSMonitoringservice@shell.com'
mail.Cc = 'anirban.a.ghosh@shell.com'
mail.Bcc = 'debashis.biswas@shell.com'
mail.Subject = 'Acknowledgement || Re: ' + FPSAlertMonitor.CTM_Mail_subject

# attachment = (os.path.expanduser(os.getenv('USERPROFILE')) + '\\Desktop\\AUTOMATION\\Logs\\loadddf.txt')
# mail.Attachments.Add(attachment)

# F = open(os.path.expanduser(os.getenv('USERPROFILE')) + '\\Desktop\\AUTOMATION\\Logs\\loadddf.txt')
# line = F.read()

(mail.body) = 'Due to cyclic Job, it will be executed successfully in the next Execution Cycle.' + (
FPSAlertMonitor.CTM_Mail_body)

mail.send

