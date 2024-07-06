import time
import os
from win32 import win32gui
import csv
import pyautogui
import subprocess

def read_file_path3(filename):
    with open(filename, "r") as file:
        lines = file.readlines()

    file_path_3 = lines[0].strip()

    return file_path_3

path0 = read_file_path3("C:\Attendance\Code\_Important.txt")
response = os.startfile(path0 + "\WNetWatcher.lnk")
print(response)

time.sleep(1)

hwnd = win32gui.FindWindow(None,"Wireless Network Watcher")

win32gui.ShowWindow(hwnd,0)

time.sleep(5)
subprocess.call("C:\Attendance\Code\dist\Core.exe")
