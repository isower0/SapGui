import win32com.client
import sys
import subprocess
import time
from datetime import datetime, date, timedelta

class SapGui():
    
    # connect QA = 550, QAS = 510, PRD = 800
    def __init__(self, path_sap,connect,username,password) -> None:
        
        self.connect = connect
        self.username = username
        self.password = password
        # r"D:\CPALL_SAP_GUI\SAPgui\saplogon.exe" 
        self.path = path_sap
        subprocess.Popen(self.path)
        time.sleep(5)
        
        self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not type(self.SapGuiAuto) == win32com.client.CDispatch:
            return

        application = self.SapGuiAuto.GetScriptingEngine
        # if self.connect == "QA":
        #     self.connection = application.OpenConnection("510 - SAP S/4 HANA Quality Assurance", True)
        # elif self.connect == "PR":
        #     self.connection = application.OpenConnection("800 - SAP S/4 HANA Production", True)
        # elif self.connect == "QAS":
        #     self.connection = application.OpenConnection("550 - SAP S/4 HANA Quality Assurance", True)
        self.connection = application.OpenConnection(connect, True)
        time.sleep(3)
        
        self.session = self.connection.Children(0)
        self.session.findById("wnd[0]").maximize
        

    def saplogon(self):
        try:
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = self.username
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = self.password
            self.session.findById("wnd[0]").sendVKey(0)
            msg = "Success Logon"
            return msg
        except:
            msg = sys.exc_info()[0]
            return msg

    def input_text(self,element: str,text):
        try:
            self.session.findById(element).text = text
            msg = "Success"
            return msg
        except:
            msg = sys.exc_info()[0]
            return msg

    def click_button(self,element: str):
        try:
            self.session.findById(element).setFocus()
            self.session.findById(element).press()
            msg = "Success"
            return msg
        except:
            msg = sys.exc_info()[0]
            return msg
            
    def sendEnter(self,key: int):
        #--------Table GUI_FKEY--------
        #https://help.sap.com/docs/sap_gui_for_windows/b47d018c3b9b45e897faf66a6c0885a8/71d8c95e9c7947ffa197523a232d8143.html?q=Table%20GUI_FKEY
        try:
            self.session.findById("wnd[0]").sendVKey(key)
            msg = "Success"
            return msg
        except:
            msg = sys.exc_info()[0]
            return msg
    
    def sendkey_anoter(self,key: int,element: str):
        #--------Table GUI_FKEY--------
        #https://help.sap.com/docs/sap_gui_for_windows/b47d018c3b9b45e897faf66a6c0885a8/71d8c95e9c7947ffa197523a232d8143.html?q=Table%20GUI_FKEY
        try:
            self.session.findById(element).sendVKey(key)
            msg = "Success"
            return msg
        except:
            msg = sys.exc_info()[0]
            return msg
    
    def check_box(self,element: str,check: bool):
        try:
            self.session.findById(element).Selected = check
            msg = "Success"
            return msg
        except:
            msg = sys.exc_info()[0]
            return msg
    
    def get_text(self,element: str):
        try:
            value = self.session.findById(element).text
            return value
        except:
            msg = sys.exc_info()[0]
            return msg

    def get_iconname(self,element: str):
        try:
            value = self.session.findById(element).IconName
            return value
        except:
            msg = sys.exc_info()[0]
            return msg
        
    def select_dropdown(self,element: str, selectkey):
        try:
            self.session.findById(element).setFocus
            self.session.findById(element).key = selectkey
            msg = "Success"
            return msg
        except:
            msg = sys.exc_info()[0]
            return msg
            
    def select_tab(self,element: str):
        try:
            self.session.findById(element).select()
            msg = "Success"
            return msg
        except:
            msg = sys.exc_info()[0]
            return msg

    def select_item(self,element: str,row,label):
        try:
            self.session.findById(element).selectItem(row,label)
            msg = "Selected"
            return msg
        except:
            msg = sys.exc_info()[0]
            return msg

    def close_window(self,element: str):
        try:
            self.session.findById(element).close
            msg = "Success"
            return msg
        except:
            msg = sys.exc_info()[0]
            return msg

    def close_sap(self):
        try:
            self.connection.CloseSession('ses[0]')
            msg = "Success"
            return msg
        except:
            msg = sys.exc_info()[0]
            return msg
