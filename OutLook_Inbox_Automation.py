#!/usr/bin/env python
import win32com.client

outlook = win32com.client.Dispatch("Outlook.application").GetNamespace("MAPI")

###"6" refers to the index of a folder - in this case, the inbox. You can change that number to reference any other folder##
inbox = outlook.GetDefaultFolder(6)

messages = inbox.Items
message = messages.GetLast()
Message_subject = message.subject
Message_body = message.body

print(message)
print(Message_subject)
print(Message_body)