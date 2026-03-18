import time
from .base_handler import BaseHandler

class IW21Handler(BaseHandler):
    def run(self, tcode, params, ev_path, mode="real"):
        self.driver.send_command("IW21")
        mapping = self.get_mapping("IW21")
        
        self.fill_generic("IW21", params)
        self.driver.send_vkey(0)
        time.sleep(0.8)

        if params.get("Tipo de nota") == "Z4":
            try:
                btn_avancar = mapping.get("POPUP_Z4", {}).get("Avançar")
                if btn_avancar:
                    self.driver.session.findById(btn_avancar).press()
                    time.sleep(0.5)
            except:
                self.driver.close_all_popups()

        self.fill_generic("IW21", params)
        return self.save_and_get_result(ev_path, mode)