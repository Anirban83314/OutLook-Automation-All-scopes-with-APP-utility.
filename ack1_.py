import win32com.client as win32

from FPSAlertMonitor import _Mail_subject , _Mail_body


def AckMailSend():
    outlook = win32.Dispatch ( 'outlook.application' )
    mail = outlook.CreateItem ( 0 )
    # mail.To = 'receipents'
    # mail.Cc = 'receipents; receipents'
    mail.Bcc = 'receipents'
    mail.Subject = 'Acknowledgement || Re: ' + _Mail_subject

    # attachment = (os.path.expanduser(os.getenv('USERPROFILE')) + '\path')
    # mail.Attachments.Add(attachment)

    # F = open(os.path.expanduser(os.getenv('USERPROFILE')) + 'path')
    # line = F.read()

    (mail.body) = 'Due to cyclic Job, it will be executed successfully in the next Execution Cycle.' + (_Mail_body)

    mail.send


if __name__ == '__main__':
    AckMailSend ( )
