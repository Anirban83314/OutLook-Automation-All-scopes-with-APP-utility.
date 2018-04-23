import sys
import os, sys, platform, easygui
import tkinter
from tkinter import *
from tkinter import messagebox, font, Frame, Scrollbar, SW, N, E, NE, SE, W, S, NW, EW, PanedWindow, Tk, scrolledtext, ttk
from subprocess import *
from threading import *
import subprocess
from subprocess import *
import subprocess
import ssmsOps.Load_OPIS_Spot_Load_Platts_Autopricing


try:
    from tkinter import *

except:
    from Tkinter import *

AppTool = Tk()
width = 400
height = 400

AppTool.resizable(1366, 768)
AppTool.configure(background='grey')
AppTool.title('OSADOL Manager :: @' + platform.node())


row = 0
while row<50:
    AppTool.rowconfigure(row, weight=1)
    AppTool.columnconfigure(row, weight=1)
    row+=1

Tab = ttk.Notebook(AppTool)
Tab.grid(row=1, column=0, columnspan=50, rowspan=49, sticky='NEWS')

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Commands Declear ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#

def disable_event():
    pass

def screenshot():

    print("Nothing configured yet !!!")

def popUp_create():
    win = tkinter.Toplevel()
    win.wm_title('Window')

    l = tkinter.Label(win, text="Input")
    l.grid(row=0, column=0)

    b = ttk.Button(win, text="Confirm", command=win.keys())
    b.grid(row=1, column=0)

    b=ttk.Button(win, text="Confirm", command=win.keys())
    b.grid(row=1, column=1)

def popUp_add_Job():
    Label(Tab5,text='Job Name').grid(row=10, column=5)
    Label(Tab5, text='Enter script PATH ~\\..').grid(row=13, column=5)

    entry1=Entry(Tab5)
    entry2=Entry(Tab5)

    entry1.grid(row=10, column=10)
    entry2.grid(row=13, column=10)

    Button ( Tab5 , text="Confirm" ).grid ( row=15 , column=8 , sticky=W , pady=4 )
    Button ( Tab5 , text="Abort" ).grid ( row=15 , column=10 , sticky=W , pady=4 )


def Exit():
    print ( 'EXIT' )
    AppTool.destroy()


#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~TAB Declearing ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
#Tab1
Tab1 = ttk.Frame(Tab)
Tab.add(Tab1, text='Jobs')
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Items under Jobs are defined~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
def  Permit_Vision():
    print('Initiaiting Permit Vision operation on every WednesDay Night @ 12:45 AM ')
    exec(open("D:\OutLook Mail Automation\PermitVisionMailAck.py").read())
    print('Operation Terminated !!')

b = Button(Tab1, text="Permit Vision", command=Permit_Vision, padx=4, pady=2, fg="#00AAAA" , bg='#000000' ).grid ( row=1, column=0, sticky='WS')


def Resource_Lock():
    print ( 'Initiaiting Resource Lock operation for both 12 PM & 9 PM daily.  ' )
    exec ( open ( "D:\OutLook Mail Automation\MailAckRsLc.py" ).read ( ) )
    print ( 'Operation Terminated !!' )

b = Button(Tab1, text="Resource Lock", command=Resource_Lock, padx=4, pady=2, fg="#00AAAA" , bg='#000000' ).grid ( row=1, column=1, sticky='WS')

def wrk_archive():
    print('Archiving WRK files to the archive directory. ')
    #exec()
    print('Script is being configured !!')

b = Button(Tab1, text="WRK Archive", command=Resource_Lock, padx=4, pady=2, fg="#00AAAA" , bg='#000000' ).grid ( row=1, column=2, sticky='WS')

def xhub_manual_monitoring():
    print ( 'When XHub FMB getting so high count mails, we can reduce the pressure using this script. ' )
    exec(open('D:\OutLook Mail Automation\XHubAutomation.py').read())
    print ( 'Operation terminated on demand !!' )

b = Button(Tab1, text="XHub Monitoring", command=xhub_manual_monitoring, padx=4, pady=2, fg="#00AAAA" , bg='#000000' ).grid ( row=1, column=3, sticky='WS')


ttk.Separator(Tab1).place(x=0, y=40, relwidth=1)
ttk.Label(Tab1, text="All Jobs").place(x=0, y=42)

#Tab2
Tab2 = ttk.Frame(Tab)
Tab.add(Tab2, text='Mails')

#Tab3
Tab3 = ttk.Frame(Tab)
Tab.add(Tab3, text='Status')

#Tab4
Tab4 = ttk.Frame(Tab)
Tab.add(Tab4, text='Process Status')

ttk.Label(Tab4, text="Current All Processes of this System").place(x=0, y=5)
ttk.Separator(Tab4).place(x=0, y=25, relwidth=15)

scrollbar = Scrollbar(Tab4)

canvas = Canvas(Tab4, height=768, width=1366, yscrollcommand=scrollbar.set)

ProcssFile = open('D:\\LOG\\taskLog.txt', 'r')

canvas.create_text(1000, 1000, text=[ProcssFile.read()], tags='text')

scrollbar.config(command=canvas.yview)
scrollbar.pack(side = RIGHT, fill = Y )

canvas.pack(side=LEFT, expand=True)

#Tab5

Tab5 = ttk.Frame(Tab)
Tab.add(Tab5, text="App Logs")

ttk.Label(Tab5, text="All Jobs related to SSMS & Control M based, can be found here!!!").place(x=0, y=5)
ttk.Separator(Tab5).place(x=0, y=25, relwidth=1)

scrollbar = Scrollbar(Tab5)


def opisIssue():
    OPIS_Autopricing_script = "D:\OutLook Mail Automation\ssmsOps\Load_OPIS_Spot_Load_Platts_Autopricing.py"

    with open('D:\\LOG\\OPIS_data.txt', 'a+') as OpisLogFile:
        OpisLogFile.write('{}'.format(subprocess.call(OPIS_Autopricing_script, shell=True)))
        OpisLogFile.close ( )

    OpisLogFileData = open('D:\\LOG\\OPIS_data.txt')


    OPISAutopricingWindow = tkinter.Toplevel()
    OPISAutopricingWindow.title('>>>> OPIS Autopricing Data Output <<<<')

    OPIS_DATA_eval = Canvas(OPISAutopricingWindow, height=768, width=1366)
    OPIS_DATA_eval.create_text(1000, 1000, text=[OpisLogFileData], tags='text')

    scrollbar = Scrollbar (OPIS_DATA_eval)
    opis_canvas = Canvas(OPIS_DATA_eval, height=1000, width=1000, yscrollcommand=scrollbar.set)

    scrollbar.config ( command=opis_canvas.yview )
    scrollbar.pack ( side=RIGHT , fill=Y )

    opis_canvas.pack(side=LEFT, expand=True)



    def recheck():
        opisDataOutput = sys.stdout
        run_opis = subprocess.call ( "ssmsOps.Load_OPIS_Spot_Load_Platts_Autopricing" , shell=True )
        OPIS_DATA = open ( 'D:\\LOG\\OPIS_data.txt' , 'w' )
        sys.stdout = opisDataOutput

        #print('Recheck query for accurate info!!')


        b = Button(OPIS_DATA_eval, command=recheck, padx=0, pady=0, fg="#00AAAA" , bg='#000000')
        b.place(x=3, y=0)


######
######
######
######
#Now find a sol for scrollable window for DB output.
#Find better resolution for OUTPUT checking details.
######
######
######
######

b = Button(Tab5, text="View OPIS Data", command=opisIssue, padx=4, pady=2, fg="#00AAAA" , bg='#000000' )
b.place(x=0, y=27)

#Tab6
Tab6 = ttk.Frame(Tab)
Tab.add(Tab6, text='++')

def add_job():
    ttk.Frame(command=popUp_add_Job())

b = Button(Tab6, text="Add Jobs", command=add_job, padx=4, pady=2, fg="#00AAAA" , bg='#000000' ).grid ( row=1, column=3, sticky='WS')

ttk.Separator(Tab6).place(x=0, y=160, relwidth=1)


def logLoc():
    print('Opening Log location!!!')
    subprocess.Popen ( 'explorer "D:\\OutLook Mail Automation\\log"' )


menubar = Menu ( AppTool , foreground="#00A7fb" , background='#000000', activebackground='#004c99', activeforeground='white' )

logmenu = Menu ( menubar , foreground="#00A7fb" , background='#000000' )
logmenu.add_command ( label='Open Log location' , command=logLoc )
logmenu.add_command ( label='Export Log' , command=logLoc )
logmenu.add_command(label='Exit', command=Exit)
menubar.add_cascade ( label='Log' , menu=logmenu )

AppTool.configure (menu=menubar)

AppTool.update()
canvas.config(scrollregion=canvas.bbox("all"))


AppTool.mainloop()
