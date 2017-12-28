# !/usr/bin/env python
import win32com.client

acknowledgement_mail = 'D:\\OutLook Mail Automation\\Templates\\ack1_BMC_DET.py'

outlook = win32com.client.Dispatch("Outlook.application").GetNamespace("MAPI")

###"6" refers to the index of a folder - in this case, the XHub_FMB. You can change that number to reference any other folder##
# But other Shared Functional Mail Boxes numeric details are often changes. So Hence those are not mentioned here.#
#This, developement is under BETA stage due to Exchange server's unstability. Cause, these numbers will be changed again & then we need to find the new numeric details#
#for each FMBs

1_FMB = outlook.Folders[9]
2_FMB_Inbox = 1_FMB.Folders[12]
3_FMB_ControlM = 2_FMB.Folders[4]

print()
print()
print()
print('/n')

# ------------------------------#
Import_Inbox_messages = 1_Inbox.Items
Import_INBOX_message = _Inbox_messages.GetFirst ( )
Secondly_Import_mail_subject = _INBOX_message.Subject
Secondly_Import__mail_body = _INBOX_message.body


# ------------------------------------#

print( "/n" )
print( )
print( )
print(  )
print()
print( "/n" )



    # ----------------------------#
ControlM_messages = FPS_FMB_ControlM.Items
CTM_DS_PROD_message = ControlM_messages.GetFirst ( )
CTM_Mail_subject = CTM_DS_PROD_message.Subject
CTM_Mail_body = CTM_DS_PROD_message.body


# --------------------------------------#

print("/n")
print()
print()
print()
print()
print()

##Hence Configuring Alert Manager.


if CTM_Mail_subject == ' test on  failed on':
    with open( 'TestMailACK.py' ) as f:
        code = compile(f.read(), "TestMailACK.py", 'exec')
        exec(code)

elif CTM_Mail_subject == 'Issue as per Req':
    with open("ackGenWCLA.py") as f:
        code = compile(f.read(), "ackGenWCLA.py", 'exec')
        exec(code)

elif CTM_Mail_subject == 'Issue as per Req':
    with open("ackGenWCLA.py") as f:
        code = compile(f.read(), "ackGenWCLA.py", 'exec')
        exec(code)

elif CTM_Mail_subject == 'Issue as per Req':
    with open("ackGenWCLA.py") as f:
        code = compile(f.read(), "ackGenWCLA.py", 'exec')
        exec(code)

elif CTM_Mail_subject == 'Issue as per Req':
    with open("ackGenWCLA.py") as f:
        code = compile(f.read(), "ackGenWCLA.py", 'exec')
        exec(code)

elif CTM_Mail_subject == 'Issue as per Req':
    with open("ackGenWCLA.py") as f:
        code = compile(f.read(), "ackGenWCLA.py", 'exec')
        exec(code)

##Hence Configuring FPS mail alert acknowledgement through this alert Manager.


else:
    print("No issues!!!")
