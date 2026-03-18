import time
import win32com.client
from pathlib import Path
from .sap_result import SapResult

class SapDriver:
    def __init__(self):
        self.session = None

    def connect(self):
        try:
            sap_gui = win32com.client.GetObject("SAPGUI")
            app = sap_gui.GetScriptingEngine
            self.session = app.Children(0).Children(0)
        except Exception as e:
            raise Exception(f"Falha ao conectar no SAP: {e}")

    def send_command(self, tcode: str):
        self.session.findById("wnd[0]/tbar[0]/okcd").text = f"/n{tcode.upper()}"
        self.session.findById("wnd[0]").sendVKey(0)

    def set_text(self, sap_id, value):
        obj = self.session.findById(sap_id)
        try:
            obj.text = str(value)
        except:
            obj.key = str(value)

    def click(self, sap_id):
        self.session.findById(sap_id).press()

    def select_tab(self, sap_id):
        self.session.findById(sap_id).select()

    def send_vkey(self, key_code: int):
        self.session.findById("wnd[0]").sendVKey(key_code)

    def get_statusbar_info(self):
        try:
            sb = self.session.findById("wnd[0]/sbar")
            return sb.MessageType, sb.Text
        except:
            return "", ""

    def popup_exists(self):
        try:
            self.session.findById("wnd[1]")
            return True
        except:
            return False

    def hardcopy(self, path: str):
        try:
            self.session.findById("wnd[0]").HardCopy(path, 2)
            return path
        except:
            return ""

    def close_all_popups(self):
        while self.popup_exists():
            try:
                self.session.findById("wnd[1]").sendVKey(0)
                time.sleep(0.5)
            except:
                break