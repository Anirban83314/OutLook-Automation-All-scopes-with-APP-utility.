#!/usr/bin/env python
import win32com.client

outlook = win32com.client.Dispatch("Outlook.application").GetNamespace("MAPI")

###"6" refers to the index of a folder - in this case, the XHub_FMB. You can change that number to reference any other folder##
# But other Shared Functional Mail Boxes numeric details are often changes. So Hence those are not mentioned here.#
#This, developement is under BETA stage due to Exchange server's unstability. Cause, these numbers will be changed again & then we need to find the new numeric details#
#for each FMBs

XHub_FMB = outlook.Folders[1]
XHub_FMB_SubFolders = XHub_FMB.Folders[10]

messages = XHub_FMB_SubFolders.Items
message = messages.GetFirst()
subject = message.Subject
Message_body = message.body
message_Count = messages.count


print(XHub_FMB)
print(XHub_FMB_SubFolders)
print(message)
print(subject)
print(Message_body)
print(message_Count)



