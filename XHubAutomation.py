#!/usr/bin/env python
import win32com.client
import datetime as dt

outlook = win32com.client.Dispatch("Outlook.application")
mapi = outlook.GetNamespace("MAPI")

date_time = dt.datetime.now()
lastQrtDateTime = dt.datetime.now() - dt.timedelta(minutes = 120)

XHub_FMB = mapi.Folders['xHub_Support_Europe SDO-DFI/14481'].Folders['Inbox']

XHub_FMB_messages = XHub_FMB.Items
XHub_FMB_messages.Sort("[ReceivedTime]", True)
XHub_FMB_message = XHub_FMB_messages.GetFirst()
mail_count = XHub_FMB_messages.count
XHub_FMB_Message_Subject = XHub_FMB_message.subject
XHub_FMB_message_Body = XHub_FMB_message.body
lastQtrMessages = XHub_FMB_messages.Restrict("[ReceivedTime] >= '" +lastQrtDateTime.strftime('%m/%d/%Y %H:%M %p')+"'")

for XHub_FMB_message in lastQtrMessages:
    print('\nMails from last 120 min\n' + XHub_FMB_Message_Subject + 'FLag 1.1')
    XHub_FMB_message = lastQtrMessages.GetFirst ( )
    XHub_FMB_message_iter = lastQtrMessages.GetNext ( )
while XHub_FMB_message:
    for XHub_FMB_message in XHub_FMB_messages:
        print(XHub_FMB_Message_Subject + 'FLag 1.2')
        if XHub_FMB_message.Unread == True:
            print ( '\n' )
            print ( XHub_FMB_Message_Subject + 'FLag 1.3' )
            XHub_FMB_message.Unread = False
            print ( '\n' )
        elif XHub_FMB_message.unread == False:
            for XHub_FMB_messages in [mail_count]:
                print ( XHub_FMB_Message_Subject + 'FLag 1.4' )
                XHub_FMB_message.Unread = False
                break
            break
print('\n No Mails pending into XHUB FMB \n')

print('End of program!!!')


#"""if XHub_FMB_message.Unread == True:
            #print ( '\n' )
            #print ( XHub_FMB_Message_Subject + 'FLag 1.3')
            #XHub_FMB_message.Unread = False
            #print('\n')
            #pass
            #if XHub_FMB_message.unread == False:
                #for XHub_FMB_messages in [mail_count]:
                    #print ( XHub_FMB_Message_Subject + 'FLag 1.4')
                    #XHub_FMB_message.Unread = False
                    #break
    #break"""
