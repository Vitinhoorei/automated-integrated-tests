import time
from .base_handler import BaseHandler

class IPHandler(BaseHandler):
    def run(self, tcode, params, ev_path, mode="real"):
        self.driver.send_command(tcode)
        mapping = self.get_mapping(tcode)
        
        self.fill_generic(tcode, params)
        self.driver.send_vkey(0)
        time.sleep(1.0)

        self.fill_generic(tcode, params)
        
        btn_lista = mapping.get("Criar lista de tarefas")
        if btn_lista:
            try:
                self.driver.session.findById(btn_lista).press()
                time.sleep(1.2)
                self.driver.close_all_popups() 
            except: pass

        self.fill_generic(tcode, params)
        self.driver.send_vkey(0)
        time.sleep(1.0)

        if tcode.upper() == "IP42":
            try:
                self.driver.session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_3400").getAbsoluteRow(0).selected = True
                self.driver.session.findById("wnd[0]/usr/btnTEXT_DRUCKTASTE_WP").press()
                time.sleep(1.0)
                self.driver.session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_3600/chkRIHSTRAT-MARK01[3,0]").selected = True
                self.driver.send_vkey(0)
            except: pass

        return self.save_and_get_result(ev_path, mode)