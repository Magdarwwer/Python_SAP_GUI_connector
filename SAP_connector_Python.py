import os
import time
import pyautogui
import subprocess

def open_SAP():
    #Start SAP Logon
    sap_gui_path = r"C:\Program Files\SAP\FrontEnd\SAPGUI\saplogon.exe" #copy path from your saplogon.exe location 
    subprocess.Popen(sap_gui_path)
    time.sleep(10)
    