import os
import time
import pyautogui
import win32com.client
import subprocess

def open_SAP():
    #Start SAP Logon
    sap_gui_path = r"C:\Program Files\SAP\FrontEnd\SAPGUI\saplogon.exe" #copy path from your saplogon.exe location 
    subprocess.Popen(sap_gui_path)
    time.sleep(10)
    
def connect_to_sap():
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        
        if application.Children.Count > 0:
            connection = application.Children(0)
            session = connection.Children(0)
            print("Connected to existing session SAP GUI")
        else:
            #Open new connection to one of the tabs available after loging into SAP
            #input the name of the SAP tab
            tab_name = "NAME_OF_YOUR_TAB" #CHANGE HERE!!!! #just copy the name you SEE, not from backend or sth. JUST the name you SEE available
            connection = application.OpenConnection(tab_name, True) 
            time.sleep(3) # wait for tab opening
            session = connection.Children(0)
            print(f"Connected to SAP GUI {tab_name}")
            
        session.findById("wnd[0]").maximize()
        return session
    except Exception as e:
        print("Couldn't connect to SAp GUI: ", e)
        return None