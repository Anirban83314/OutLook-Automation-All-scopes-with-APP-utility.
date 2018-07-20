#!/usr/bin/env python
import os
from time import strftime
from selenium import webdriver

FilePath = 'C:\\Users\\Debashis.Biswas\\Desktop\\WebModules\\'
CurrentTime = strftime("%y%m%d%H%M%S")
File_NAME = (CurrentTime + "WebMod.png")

driver = webdriver.Chrome('D:\\Python36-32\\chromedriver.exe')

driver.get("<<Some URL is here >>")
driver.maximize_window()

FilePath = 'C:\\Users\\Debashis.Biswas\\Desktop\\WebModules\\'
os.chdir (FilePath)

mailObj = driver.save_screenshot(os.path.join(FilePath + File_NAME))
#time.sleep(10)

driver.quit()
##Checking Newest File for acknowledgement.
FilePath = 'C:\\Users\\Debashis.Biswas\\Desktop\\WebModules\\'
os.chdir (FilePath)

filesNew = sorted ( os.listdir ( os.getcwd () ) , key=os.path.getmtime )

# oldest = filesNew[0]
newest = filesNew[-1]

# print "Oldest:", oldest
print ( "Newest:" , newest )
print ( "All by modified oldest to newest:" , filesNew )
