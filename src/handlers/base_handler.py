import yaml
import time
from src.core.sap_result import SapResult

class BaseHandler:
    def __init__(self, driver, field_map_path="configs/field_map.yaml"):
        self.driver = driver
        with open(field_map_path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
        self.field_map = {str(k).upper(): v for k, v in data.items()}

    def get_mapping(self, tcode):
        t_map = self.field_map.get(tcode.upper(), {})
        
        flat_map = {}
        for screen in t_map.values():
            if isinstance(screen, dict):
                flat_map.update(screen)
        return flat_map

    def fill_generic(self, tcode, params):
        """Preenche campos baseados no mapeamento YAML."""
        mapping = self.get_mapping(tcode)
        remaining = {}
        
        for key, value in params.items():
            sap_id = mapping.get(key)
            if sap_id and value:
                try:
                    self.driver.set_text(sap_id, value)
                except:
                    remaining[key] = value
            else:
                remaining[key] = value
        return remaining

    def save_and_get_result(self, ev_path, mode="real"):
        """Lógica de salvamento e captura de status bar."""
        if mode == "real":
            self.driver.send_vkey(11)
            time.sleep(1.5)
            self.driver.close_all_popups()
        else:
            self.driver.send_vkey(0) 
            time.sleep(0.5)

        m_type, m_text = self.driver.get_statusbar_info()
        status = "PASS" if m_type not in ["E", "A", "X"] else "FAIL"
        source = "STATUSBAR" if status == "FAIL" else "OK"
        
        final_ev = self.driver.hardcopy(ev_path)
        return SapResult(status, source, m_text or "Executado", final_ev)