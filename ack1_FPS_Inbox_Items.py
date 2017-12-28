import win32com.client as win32

from FPSAlertMonitor import FPS_Mail_subject , FPS_Mail_body


def AckMailSend():
    outlook = win32.Dispatch ( 'outlook.application' )
    mail = outlook.CreateItem ( 0 )
    # mail.To = ''
    # mail.Cc = 'GXSOMWIPRODSMonitoringservice@shell.com; anirban.a.ghosh@shell.com'
    mail.Bcc = 'debashis.biswas@shell.com'
    mail.Subject = 'Acknowledgement || Re: ' + FPS_Mail_subject

    # attachment = (os.path.expanduser(os.getenv('USERPROFILE')) + '\\Desktop\\AUTOMATION\\Logs\\loadddf.txt')
    # mail.Attachments.Add(attachment)

    # F = open(os.path.expanduser(os.getenv('USERPROFILE')) + '\\Desktop\\AUTOMATION\\Logs\\loadddf.txt')
    # line = F.read()

    (mail.body) = 'Due to cyclic Job, it will be executed successfully in the next Execution Cycle.' + (FPS_Mail_body)

    mail.send


if __name__ == '__main__':
    AckMailSend ( )
