import win32com.client as win32

import MasterAlertManager



outlook = win32.Dispatch ( 'outlook.application' )
mail = outlook.CreateItem ( 0 )
mail.To = 'recepients'
mail.Cc = 'recepients'
mail.Bcc = 'recepients'
mail.Subject = 'Acknowledgement || Re: ' + MasterAlertManager._Mail_subject

# attachment = (os.path.expanduser(os.getenv('USERPROFILE')) + '\\Path\\to\\the\\Attachment\\')
# mail.Attachments.Add(attachment)

# F = open(os.path.expanduser(os.getenv('USERPROFILE')) + '\\Path\\to\\the\\Attachment\\')
# line = F.read()

(mail.body) = 'Small_description.' + (
MasterAlertManager._Mail_body)

mail.send

