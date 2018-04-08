import os, sys, platform, easygui
from tkinter import *
from tkinter import messagebox, font, Frame, Scrollbar, SW, N, E, NE, SE, W, S, NW, EW, PanedWindow, Tk, scrolledtext, ttk
from subprocess import *

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


#Tab2
Tab2 = ttk.Frame(Tab)
Tab.add(Tab2, text='Mails')

#Tab3
Tab3 = ttk.Frame(Tab)
Tab.add(Tab3, text='Status')

#Tab4
Tab4 = ttk.Frame(Tab)
Tab.add(Tab4, text='Process Status')

#Tab5
Tab5 = ttk.Frame(Tab)
Tab.add(Tab5, text='++')


appPane = PanedWindow(orient=HORIZONTAL)
#appPane.pack()

top = Label(appPane, text='TOP PANE')
appPane.add(top, stretch="always")



AppTool.mainloop()
