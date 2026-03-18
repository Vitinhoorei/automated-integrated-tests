import time
from .base_handler import BaseHandler
from src.core.sap_result import SapResult

class IW31Handler(BaseHandler):
    def run(self, tcode, params, ev_path, mode="real"):
        self.driver.send_command(tcode)
        mapping = self.get_mapping(tcode)
        rem = self.fill_generic(tcode, params)
        self.driver.send_vkey(0) 
        time.sleep(1.0)

        if params.get("Criar Nota") == "X":
            btn_nota = mapping.get("Criar Nota")
            if btn_nota:
                try:
                    self.driver.session.findById(btn_nota).press()
                    time.sleep(0.8)
                except: pass

        if params.get("Aba Operações") == "X":
            tab_id = mapping.get("Aba Operações")
            if tab_id:
                try:
                    self.driver.session.findById(tab_id).select()
                    time.sleep(0.6)
                except: pass

        self.fill_generic(tcode, rem)
        
        return self.save_and_get_result(ev_path, mode)