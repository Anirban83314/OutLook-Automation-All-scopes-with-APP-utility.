import os
import tkinter
from tkinter import font, messagebox, Button, Frame, PanedWindow, Widget, scrolledtext, Tk, SW, LEFT, RIGHT, Label
from collections import defaultdict
import itertools

try:
    from tkinter import *
except ImportError:
    from Tkinter import *

ManagerTool = tkinter.Tk()

width = 400
hight = 400

ws = ManagerTool.winfo_screenwidth()
hs = ManagerTool.winfo_screenmmheight()
x = (ws/1) - (width/1)
y = (hs/1) - (hight/1)

ManagerTool.geometry('%dx%d+%d+%d' % (width, hight, x, y))

ManagerTool.resizable(1366,768)
ManagerTool.configure(background='white')
ManagerTool.lift()
ManagerTool.title('OSADOL MANAGER')
ManagerTool.rowconfigure(0, weight=1)
ManagerTool.rowconfigure(1, weight=1)

def __init__(self):
    Tk.Tk.__init__ ( self )
    self.title ( "Handling WM_DELETE_WINDOW protocol" )
    self.geometry ( "500x300+500+200" )
    self.make_topmost ()
    self.protocol ( "WM_DELETE_WINDOW", self.on_exit )

def on_Exit(self):
    """When you click to exit, this function is called"""
    if messagebox.askyesno ( "Exit", "Do you want to quit the application?" ):
        self.destroy ()

def make_topmost(self):
    """Makes this window the topmost window"""
    self.lift ()
    self.attributes ( "-topmost", 1 )
    self.attributes ( "-topmost", 0 )

def Services():
    print('Services')
    os.chdir ( 'C:\\Program Files\\Microsoft Office\\root\\Office16\\' )
    os.startfile (r"C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE")
b = Button ( ManagerTool, text="OutLook", command=Services, padx=6 )
b.place(anchor=SW)
b.pack (side = LEFT)

def Exit():
    print('EXIT')
    ManagerTool.destroy()
b = Button ( ManagerTool, text="Exit", command=Exit, padx=3)
b.place(anchor=SW)
b.pack (side = LEFT)


ManagerTool.mainloop()