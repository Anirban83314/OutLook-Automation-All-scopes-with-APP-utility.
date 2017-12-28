# !/usr/bin/env python
import win32com.client

acknowledgement_mail = 'D:\\OutLook Mail Automation\\Templates\\ack1_CTM_DS_PRD.py'

outlook = win32com.client.Dispatch("Outlook.application").GetNamespace("MAPI")

###"6" refers to the index of a folder - in this case, the XHub_FMB. You can change that number to reference any other folder##
# But other Shared Functional Mail Boxes numeric details are often changes. So Hence those are not mentioned here.#
#This, developement is under BETA stage due to Exchange server's unstability. Cause, these numbers will be changed again & then we need to find the new numeric details#
#for each FMBs

FPS_FMB = outlook.Folders[9]
FPS_FMB_Inbox = FPS_FMB.Folders[12]
FPS_FMB_ControlM = FPS_FMB.Folders[4]

print(FPS_FMB)
print(FPS_FMB_Inbox)
print(FPS_FMB_ControlM)
print('/n')
print('>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>Checking up alert manager configuration ## Optional ## <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<')


# Checking FPS inbox for alerts#
# ------------------------------#
FPS_Inbox_messages = FPS_FMB_Inbox.Items
FPS_INBOX_message = FPS_Inbox_messages.GetFirst ( )
f_p_s__mail_subject = FPS_INBOX_message.Subject
f_p_s__mail_body = FPS_INBOX_message.body

# Now checking parameter for FPS INBOX#
# ------------------------------------#

print( "/n" )
print( FPS_INBOX_message )
print( f_p_s__mail_subject )
print( f_p_s__mail_body )
print( '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> FPS INBOX ends here. <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<' )
print( "/n" )


# Checing Control M for alerts#
    # ----------------------------#
ControlM_messages = FPS_FMB_ControlM.Items
CTM_DS_PROD_message = ControlM_messages.GetFirst ( )
CTM_Mail_subject = CTM_DS_PROD_message.Subject
CTM_Mail_body = CTM_DS_PROD_message.body

# Now checking parameter for CTM_DS_PROD#
# --------------------------------------#

print("/n")
print(CTM_DS_PROD_message)
print(CTM_Mail_subject)
print(CTM_Mail_body)
print('>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>CTM_DS_PROD ends here.<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<')
print("/n")

#def CTM_Mail_SubjectItems():
    #return CTM_Mail_subject

#def CTM_Mail_BodyItems():
    #return CTM_Mail_body

##Hence Configuring ControlM mail alert management through Alert Manager.


if CTM_Mail_subject == 'FPS Price POLL alive test on  failed on www.fps.shell.com':
    with open( 'PollAliveTestMailACK.py' ) as f:
        code = compile(f.read(), "PollAliveTestMailACK.py", 'exec')
        exec(code)

elif CTM_Mail_subject == 'FPS_CH: File Watcher Job Failed':
    with open("ackGenWCLA.py") as f:
        code = compile(f.read(), "ackGenWCLA.py", 'exec')
        exec(code)

elif CTM_Mail_subject == 'FPS_AR: File Watcher Job Failed':
    with open("ackGenWCLA.py") as f:
        code = compile(f.read(), "ackGenWCLA.py", 'exec')
        exec(code)

elif CTM_Mail_subject == 'FPS_PL: File watcher failed for PL volume import':
    with open("ackGenWCLA.py") as f:
        code = compile(f.read(), "ackGenWCLA.py", 'exec')
        exec(code)

elif CTM_Mail_subject == 'FPS_AT: Check Log Failed':
    with open("ackGenWCLA.py") as f:
        code = compile(f.read(), "ackGenWCLA.py", 'exec')
        exec(code)

elif CTM_Mail_subject == 'FPS_BG: Error in Importing the Effective-Prices':
    with open("ackGenWCLA.py") as f:
        code = compile(f.read(), "ackGenWCLA.py", 'exec')
        exec(code)

##Hence Configuring FPS mail alert acknowledgement through this alert Manager.


else:
    print("No issues!!!")