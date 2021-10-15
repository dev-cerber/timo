# import time
# from PIL import ImageGrab
#
# img = ImageGrab.grab(bbox=(100, 200, 100, 200))
# img.save('1.png')

from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import *
import win32gui
import sys

hwnd = win32gui.FindWindow(None, 'C:\Windows\system32\cmd.exe')
app = QApplication(sys.argv)
screen = QApplication.primaryScreen()
img = screen.grabWindow(hwnd).toImage()
img.save("screenshot.jpg")