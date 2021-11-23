import win32com.client
# import datetime
#import time
from tkinter import *
from tkinter import messagebox as mb
#import os
import webbrowser
import eel
# import rastr_calc_kor_start
# from rastr_calc_kor_start import start

eel.init('web')
#
# def my_other_thread():
#     while True:
#         # print("I'm a thread")
#         eel.sleep(1.0) # Use eel.sleep(), not time.sleep()
#
# eel.spawn(my_other_thread)

eel.start('main.html' ) # Don't block on this call

# while True:
#      # print("I'm a main loop")
#      eel.sleep(1.0) # Use eel.sleep(), not time.sleep()

def print_num(n):
    print('Got this from Javascript:', n)

# Call Javascript function, and pass explicit callback function
eel.js_gost()(print_num)

# Do the same with an inline lambda as callback
eel.js_gost()(lambda n: print('Got this from Javascript:', n))


mb.showinfo("Инфо","88888888")

GL = None
LogFile = None
rastr = win32com.client.Dispatch("Astra.Rastr")
visual_set = 0  # 1 задание через GUI, 0  - в коде

# if visual_set == 0:
#     start()

# webbrowser.open("LogFile.txt")