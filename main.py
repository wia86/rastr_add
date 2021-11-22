import win32com.client
# import datetime
#import time
#from tkinter import *
#from tkinter import messagebox as mb
#import os
import webbrowser
import eel
import rastr_calc_kor_start
# from rastr_calc_kor_start import start


GL = None
LogFile = None
rastr = win32com.client.Dispatch("Astra.Rastr")
visual_set = 0  # 1 задание через GUI, 0  - в коде

# if visual_set == 0:
#     start()

# webbrowser.open("LogFile.txt")