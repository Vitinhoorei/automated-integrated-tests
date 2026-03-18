import time
from .base_handler import BaseHandler

class IW32Handler(BaseHandler):
    def run(self, tcode, params, ev_path, mode="real"):
        self.driver.send_command("IW32")
        mapping = self.get_mapping("IW32")
        ordem_id = mapping.get("Ordem")
        if ordem_id:
            self.driver.set_text(ordem_id, params.get("Ordem"))
        
        self.driver.send_vkey(0) 
        time.sleep(1.0)
        
        self.fill_generic("IW32", params)
        
        return self.save_and_get_result(ev_path, mode)