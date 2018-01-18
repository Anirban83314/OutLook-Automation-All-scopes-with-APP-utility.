import os
import tkinter
from tkinter import messagebox, font, Frame, Scrollbar, SW, N, E, NE, SE, W, S, NW, EW, PanedWindow, Tk
from subprocess import *

try:
    from tkinter import *

except:
    from Tkinter import *


ManagerTool = Tk()

width = 400
hight = 400

ws = ManagerTool.winfo_screenwidth()
hs = ManagerTool.winfo_screenmmheight()
x = (ws / 1) - (width / 1)
y = (hs / 1) - (hight / 1)

ManagerTool.geometry('%dx%d+%d+%d' % (width, hight, x, y))

ManagerTool.resizable(1366, 768)
ManagerTool.configure(background='white')
ManagerTool.lift()
ManagerTool.title('OSADOL MANAGER')
ManagerTool.rowconfigure(0, weight=1)
ManagerTool.rowconfigure(1, weight=1)

def __init__(self):
    Tk.Tk.__init__(self)
    self.title("Handling WM_DELETE_WINDOW protocol")
    self.geometry("500x300+500+200")
    self.make_topmost()
    self.protocol("WM_DELETE_WINDOW", self.on_exit)

def on_Exit(self):
    """When you click to exit, this function is called"""
    if messagebox.askyesno("Exit", "Do you want to quit the application?"):
        self.destroy()

def make_topmost(self):
    """Makes this window the topmost window"""
    self.lift()
    self.attributes("-topmost", 1)
    self.attributes("-topmost", 0)
##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Button Area defined~~~~~~~~~~~~~~~~~~~~~~~~~~~~##
def Services():
    print('Services')
    os.chdir('C:\\Program Files\\Microsoft Office\\root\\Office16\\')
    os.startfile(r"C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE")

b = Button(ManagerTool, text="OutLook", command=Services, padx=6)
b.place(anchor=SW)
b.pack(side=LEFT)

def Exit():
    print('EXIT')
    ManagerTool.destroy()

b = Button(ManagerTool, text="Exit", command=Exit, padx=3)
b.place(anchor=SW)
b.pack(side=LEFT)


##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Menubar Commands~~~~~~~~~~~~~~~~~~~~~~~~~~~~~##
def log():
    print('Log command has not defined yet.')


def user():
    print('User detals will be shown here. ')


def scheduler():
    print('Job scheduler is being showing here .')


def apps():
    print('App details. which will show running scripts & there locations. but not read or write availablity. ')


def share():
    print('Share Logs & app location along with last run status with time stamp to avoide confusion. ')


##~~~~~~~~~~~~~~~~~~~~~~~~~~~~Menubar commands finished here~~~~~~~~~~~~~~~~~~~~~~##

##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Menu Bar items defined here~~~~~~~~~~~~~~~~~~~~~~~~##

menubar = Menu(ManagerTool)
menubar.add_command(label='Log', command=log)
menubar.add_command(label='User', command=user)
menubar.add_command(label='Seheduler', command=scheduler)
menubar.add_command(label='Apps', command=apps)
menubar.add_command(label='Share', command=share)
ManagerTool.configure(menu=menubar)
##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~End of Menu Bar items~~~~~~~~~~~~~~~~~~~~~~~~~~~##

##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~Pane window define here~~~~~~~~~~~~~~~~~~~~~~~~~~~##
m=PanedWindow(orient=VERTICAL)
m.pack(fill=BOTH,expand=1)

top = Label(m, text="Top Pane")
m.add(top)

buttom = Label(m, text="Buttom pane")
m.add(buttom)

##~~~~~~~~~~~~~~~~~~~~~~~~~~~~~End of Pane Window here~~~~~~~~~~~~~~~~~~~~~~~~~~~~##
ManagerTool.mainloop()
