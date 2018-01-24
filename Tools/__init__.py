from tkinter.scrolledtext import ScrolledText

__project__ = 'OSADOL.'
__author__ = 'Jeys_Ozzius.'
__descript__ = "Copyright (c) 2017 Anirban Ghosh(Anirban83314) & Debashis Biswas(deb991)."
__URL__ = 'https://github.com/Anirban83314/OutLook-Automation-All-scopes-with-APP-utility.'
__NB__ = 'For more information, please see github page & all commit details.'
__CipherSig__ = 'This project & all associate files are encrypted under PBEncryption cryptography. Another details will be available at the end of this pgogram.'

import os
import tkinter
from tkinter import messagebox , font , Frame , Scrollbar , SW , N , E , NE , SE , W , S , NW , EW , PanedWindow , Tk, scrolledtext
from subprocess import *
import platform

try:
    from tkinter import *

except:
    from Tkinter import *

ManagerTool = Tk ( )
width = 400
hight = 400

ws = ManagerTool.winfo_screenwidth ( )
hs = ManagerTool.winfo_screenmmheight ( )
x = (ws / 1) - (width / 1)
y = (hs / 1) - (hight / 1)

ManagerTool.geometry ( '%dx%d+%d+%d' % (width , hight , x , y) )

ManagerTool.resizable ( 1366 , 768 )
ManagerTool.configure ( background='white' )
ManagerTool.lift ( )
ManagerTool.title ( 'OSADOL MANAGER' )
ManagerTool.rowconfigure ( 0 , weight=1 )
ManagerTool.rowconfigure ( 1 , weight=1 )


def __init__(self):
    Tk.Tk.__init__ ( self )
    self.title ( "Handling WM_DELETE_WINDOW protocol" )
    self.geometry ( "500x300+500+200" )
    self.make_topmost ( )
    self.protocol ( "WM_DELETE_WINDOW" , self.on_exit )

def on_Exit(self):
    """When you click to exit, this function is called"""
    if messagebox.askyesno ( "Exit" , "Do you want to quit the application?" ):
        self.destroy()

def make_topmost(self):
    """Makes this window the topmost window"""
    self.lift ( )
    self.attributes ( "-topmost" , 1 )
    self.attributes ( "-topmost" , 0 )

    ##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Button Area defined~~~~~~~~~~~~~~~~~~~~~~~~~~~~##
def Jobs():
    print ( 'Services' )
    os.chdir ( 'C:\\Program Files\\Microsoft Office\\root\\Office16\\' )
    os.startfile ( r"C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE" )

b = Button ( ManagerTool , text="Jobs" , command=Jobs , padx=4 , fg="#00A7fb" , bg='#000000' ).grid ( row=0, column=0, sticky='NW' )



def Mails(self):
    print ( 'Job details ' )

b = Button ( ManagerTool , text="Mails" , command=Mails , padx=4 , fg="#00A7fb" , bg='#000000' ).grid ( row=0 ,column=1, sticky='NW' )

def status(self):
    print ( 'Rpinting Process status & job status on the banner window!!!' )

b = Button ( ManagerTool , text="Status" , command=status , padx=4 , fg="#00A7fb" , bg='#000000' ).grid ( row=0, column=2 ,sticky='NW' )


def process_status(self):
    print ( 'Showing only running process status detasil. Else nothing will return as per job scheduling. ' )

b = Button ( ManagerTool , text="Process Status" , command=process_status , padx=4 , fg="#00A7fb" ,
                 bg='#000000' ).grid ( row=0 , column=3 , sticky='NW' )

def Exit(self):
    print ( 'EXIT' )
    ManagerTool.destroy ( )

b = Button ( ManagerTool , text="Exit" , command=Exit , padx=4 , fg="#00AAAA" , bg='#000000' ).grid ( row=0 ,
                                                                                                          column=4 ,
                                                                                                          sticky='NW' )

def Test(self):
    print ( 'Testing Manager during BETA stage!!!' )
    print ( platform.platform ( ) )
    print ( platform.machine ( ) )
    print ( platform.uname ( ) )

##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Menubar Commands~~~~~~~~~~~~~~~~~~~~~~~~~~~~~##
def assending_a2z(self):
    print ( 'Showing logs from assending order. But as per Current program. !!!!' )
    print ( 'Log command has not defined yet.' )

def user(self):
    print ( 'User detals will be shown here. ' )

def scheduler(self):
    print ( 'Job scheduler is being showing here .' )

def apps(self):
    print ( 'App details. which will show running scripts & there locations. but not read or write availablity. ' )

def share(self):
    print ( 'Share Logs & app location along with last run status with time stamp to avoide confusion. ' )

def Help(self):
    print ( open ( 'https://github.com/Anirban83314/OutLook-Automation-All-scopes-with-APP-utility.' ) )

def About(self):
    print ( open ('https://github.com/Anirban83314/OutLook-Automation-All-scopes-with-APP-utility./blob/master/LICENSE' ) )

    ##~~~~~~~~~~~~~~~~~~~~~~~~~~~~Menubar commands finished here~~~~~~~~~~~~~~~~~~~~~~##

    ##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Menu Bar items defined here~~~~~~~~~~~~~~~~~~~~~~~~##

menubar = Menu ( ManagerTool , foreground="#00A7fb" , background='#000000' )

logmenu = Menu ( menubar , foreground="#00A7fb" , background='#000000' )
logmenu.add_command ( label='A-Z view' , command=assending_a2z )
logmenu.add_command ( label='Z-A view' , command=assending_a2z )
logmenu.add_command ( label='Open Log location' , command=assending_a2z )
logmenu.add_command ( label='Export Log' , command=assending_a2z )
logmenu.add_command ( label='Exit' , command=Exit )
menubar.add_cascade ( label='Log' , menu=logmenu )

usermenu = Menu ( menubar , foreground="#00A7fb" , background='#000000' )
usermenu.add_command ( label='User Details' , command=user )
menubar.add_cascade ( label='User' , menu=usermenu )

schedulermenu = Menu ( menubar , foreground="#00A7fb" , background='#000000' )
schedulermenu.add_command ( label='View Scheduled Jobs' , command=scheduler )
schedulermenu.add_command ( label='View all scheduled jobs' , command=scheduler )
schedulermenu.add_command ( label='All activities of scheduled jobs' , command=scheduler )
menubar.add_cascade ( label='Scheduler' , menu=schedulermenu )

Appmenu = Menu ( menubar , foreground="#00A7fb" , background='#000000' )
Appmenu.add_command ( label='Control M Wrapper[Beta 1.0.0.1]' , command=apps )
menubar.add_cascade ( label='Appmenu' , menu=Appmenu )

Sharemenu = Menu ( menubar , foreground="#00A7fb" , background='#000000' )
Sharemenu.add_command ( label='Share Log details' , command=share )
menubar.add_cascade ( label='Share' , menu=Sharemenu )

HelpAbout = Menu ( menubar , foreground="#00A7fb" , background='#000000' )
HelpAbout.add_command ( label='Help' , command=Help )
HelpAbout.add_command ( label='About' , command=About )
menubar.add_cascade ( label='Help' , menu=HelpAbout )

ManagerTool.configure ( menu=menubar )
##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~End of Menu Bar items~~~~~~~~~~~~~~~~~~~~~~~~~~~##

##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Pane window define here~~~~~~~~~~~~~~~~~~~~~~~~~~~##
m = PanedWindow ( orient=VERTICAL ).grid ( row=0 , column=0 , sticky=W )

top = Label ( m , text="Prime Jobs" ).grid ( row=1 , column=3 , sticky=NW )

buttom = Label ( m , text="All Jobs List" ).grid ( row=2 , column=3 , sticky=W )



##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~End of Pane Window here~~~~~~~~~~~~~~~~~~~~~~~~~~~~##

ManagerTool.mainloop()
