#!/usr/bin/env python
import os
from PIL import ImageGrab
import time
from time import strftime

ScreenShot = os.getenv('USERPROFILE') + '\\Desktop\\Screenshots\\'
CurrentTime = strftime("%y%m%d%H%M%S")
time.sleep(1)
ImageGrab.grab().save(ScreenShot + CurrentTime + "Scrsht.jpeg", "jpeg")


