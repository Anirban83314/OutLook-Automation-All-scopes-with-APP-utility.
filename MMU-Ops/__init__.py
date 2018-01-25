__project__ = 'OSADOL.'
__author__ = 'Jeys_Ozzius.'
__descript__ = "Copyright (c) 2017 Anirban Ghosh(Anirban83314) & Debashis Biswas(deb991)."
__URL__ = 'https://github.com/Anirban83314/OutLook-Automation-All-scopes-with-APP-utility.'
__NB__ = 'For more information, please see github page & all commit details.'
__CipherSig__ = 'This project & all associate files are encrypted under PBEncryption cryptography. Another details will be available at the end of this pgogram.'

import os
import sys
import subprocess
from ctypes import windll
import string
import mmap

def get_drives():
    drives = []
    bitmask = windll.kernel32.GetLogicalDrives()
    for letter in string.ascii_uppercase:
        if bitmask & 1:
            drives.append(letter)
        bitmask >>= 1

    return drives

def check_indexing():
    print('Indexing whole drive of the system.')
    directory = ('C:\\')
    baseDir = "\\Outlook automation and app utility"
    try:
        for root, folders, files in os.walk(directory):
            with open('index.txt', 'a') as f:
                print(root, file=f)
                #folders\\ + files, file=f)

    except FileNotFoundError:
        print('[WinError 3] The system cannot find the path specified:')

def find_parent_locals():
    print('For this Operation Python 3.6 required !!')
    Prime_dependency = 'Python36-32'
    parent_locals = 'OutLook-Automation-All-scopes-with-APP-utility.'
    with open('index.txt', 'rb', 0) as file, \
        mmap.mmap(file.fileno(), 0, access=mmap.ACCESS_READ) as s:
        if s.find(b'OutLook-Automation-All-scopes-with-APP-utility') != -1:
            print('True')
        else:
            print('Exception!!! ')
get_drives()
check_indexing()
find_parent_locals()
