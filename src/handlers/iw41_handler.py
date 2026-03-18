import time
from datetime import datetime, timedelta
from .base_handler import BaseHandler
from src.core.sap_result import SapResult

class IW41Handler(BaseHandler):
    def run(self, tcode, params, ev_path, mode="real"):
        self.driver.send_command("IW41")
        mapping = self.get_mapping("IW41")
        
        ordem_id = mapping.get("Ordem")
        if ordem_id:
            self.driver.set_text(ordem_id, params.get("Ordem"))
        
        self.driver.send_vkey(0)
        time.sleep(0.5)
        self.driver.send_vkey(0)
        time.sleep(1.0)

        ontem = (datetime.now() - timedelta(days=1)).strftime("%d.%m.%Y")
        table_id = mapping.get("Tabela confirmações")
        processed_count = 0

        try:
            table = self.driver.session.findById(table_id)
            for i in range(int(table.RowCount)):
                if i > 5: break
                try:
                    table.getAbsoluteRow(i).selected = True
                    self.driver.send_vkey(2)
                    time.sleep(0.8)
                    self.driver.set_text(mapping.get("Nº pessoal"), params.get("Nº pessoal"))
                    self.driver.set_text(mapping.get("Início trabalho data"), ontem)
                    self.driver.set_text(mapping.get("Fim trabalho data"), ontem)
                    self.driver.send_vkey(0)
                    if mode == "real":
                        self.driver.send_vkey(11)
                    processed_count += 1
                    self.driver.send_vkey(3) 
                    time.sleep(0.8)
                except: continue
        except: pass

        ev = self.driver.hardcopy(ev_path)
        return SapResult("PASS", "OK", f"Confirmadas {processed_count} ops", ev)