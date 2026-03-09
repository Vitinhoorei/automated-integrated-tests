from __future__ import annotations

import win32com.client
import yaml
import time
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Optional
from sap_screen_dump import dump_screen

# MODELOS
@dataclass
class SapResult:
    status: str
    source: str
    message: str
    evidence_path: str = ""

class SapAutomation:
    """
    Conecta a uma sessão SAP GUI já aberta e logada.
    Executa transações, preenche campos inteligentemente navegando por múltiplas telas.
    """

    def __init__(self, field_map_path: str = "configs/field_map.yaml"):
        self.session = None
        self.field_map = self._load_field_map(field_map_path)

    def connect_existing_session(self) -> None:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine
        connection = application.Children(0)
        self.session = connection.Children(0)

    def _ensure_session(self) -> None:
        if self.session is None:
            self.connect_existing_session()

    def _load_field_map(self, path: str) -> dict:
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = yaml.safe_load(f) or {}
            return {str(k).upper(): (v or {}) for k, v in data.items()}
        except FileNotFoundError:
            return {}

    @staticmethod
    def _norm_key(text: str) -> str:
        text = (text or "").strip().lower()
        text = unicodedata.normalize("NFD", text)
        text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
        return " ".join(text.split())

    def _screen_key(self) -> str:
        try:
            prog = str(self.session.Info.Program).strip()
            scr = str(self.session.Info.ScreenNumber).strip()
            if scr.isdigit() and len(scr) < 4:
                scr = scr.zfill(4)
            return f"{prog}|{scr}"
        except Exception:
            return "UNKNOWN|UNKNOWN"

    def _statusbar_text(self) -> str:
        try:
            return self.session.findById("wnd[0]/sbar").Text or ""
        except Exception:
            return ""

    def _statusbar_type(self) -> str:
        try:
            return (self.session.findById("wnd[0]/sbar").MessageType or "").strip().upper()
        except Exception:
            return ""

    def _popup_exists(self) -> bool:
        try:
            self.session.findById("wnd[1]")
            return True
        except Exception:
            return False

    def _popup_text(self) -> str:
        try:
            wnd1 = self.session.findById("wnd[1]")

            for name in ("txtMESSTXT1", "txtSPOP-TEXTLINE1"):
                try:
                    return wnd1.findByName(name, "GuiTextField").Text
                except Exception:
                    pass

            texts =[]
            try:
                for i in range(wnd1.Children.Count):
                    child = wnd1.Children(i)
                    txt = (getattr(child, "Text", "") or "").strip()
                    if txt:
                        texts.append(txt)
            except Exception:
                pass

            return " | ".join(texts) if texts else (getattr(wnd1, "Text", "") or "Popup detectado")
        except Exception:
            return ""

    def _dismiss_popup(self) -> None:
        """Tenta fechar o popup de forma inteligente clicando em Sim ou Continuar."""
        try:
            wnd1 = self.session.findById("wnd[1]")
        except Exception:
            return

        botoes_confirmacao = [
            "tbar[0]/btn[0]",        
            "usr/btnBUTTON_1",       
            "tbar[0]/btn[11]",       
        ]

        for btn in botoes_confirmacao:
            try:
                wnd1.findById(btn).press()
                time.sleep(0.4)
                return
            except Exception:
                continue

        try:
            wnd1.sendVKey(0)
            time.sleep(0.4)
        except Exception:
            pass

    def _hardcopy_wnd0(self, out_path: str) -> str:
        try:
            target = Path(out_path)
            target.parent.mkdir(parents=True, exist_ok=True)

            bmp = target.with_suffix(".bmp")
            self.session.findById("wnd[0]").HardCopy(str(bmp), 2)

            if bmp.suffix.lower() == target.suffix.lower():
                return str(bmp)

            from PIL import Image
            with Image.open(bmp) as img:
                img.save(target)

            try:
                bmp.unlink()
            except Exception:
                pass

            return str(target)
        except Exception:
            return ""

    def _capture_object_region(self, obj_id: str, out_path: str, pad: int = 12) -> bool:
        try:
            from PIL import ImageGrab
            obj = self.session.findById(obj_id)

            left = int(getattr(obj, "ScreenLeft", 0))
            top = int(getattr(obj, "ScreenTop", 0))
            width = int(getattr(obj, "Width", 0))
            height = int(getattr(obj, "Height", 0))

            if width <= 0 or height <= 0:
                return False

            bbox = (
                max(0, left - pad),
                max(0, top - pad),
                left + width + pad,
                top + height + pad,
            )

            Path(out_path).parent.mkdir(parents=True, exist_ok=True)
            img = ImageGrab.grab(bbox=bbox)
            img.save(out_path)
            return True
        except Exception:
            return False

    def _capture_field_crop_from_hardcopy(self, field_id: str, out_path: str, pad: int = 12) -> str:
        if not field_id or not out_path:
            return ""

        base = self._hardcopy_wnd0(out_path)
        if not base:
            return ""

        try:
            from PIL import Image

            wnd0 = self.session.findById("wnd[0]")
            field = self.session.findById(field_id)

            with Image.open(base) as img:
                iw, ih = img.size

                wx, wy = wnd0.ScreenLeft, wnd0.ScreenTop
                ww, wh = wnd0.Width, wnd0.Height

                fx, fy = field.ScreenLeft, field.ScreenTop
                fw, fh = field.Width, field.Height

                x1 = int(((fx - wx) / ww) * iw) - pad
                y1 = int(((fy - wy) / wh) * ih) - pad
                x2 = int(((fx + fw - wx) / ww) * iw) + pad
                y2 = int(((fy + fh - wy) / wh) * ih) + pad

                crop = img.crop((max(0, x1), max(0, y1), min(iw, x2), min(ih, y2)))
                Path(out_path).parent.mkdir(parents=True, exist_ok=True)
                crop.save(out_path)

            return out_path
        except Exception:
            return ""

    def _capture_error_evidence(self, out_path: str, source: str, field_id: str = "") -> str:
        if not out_path:
            return ""

        src = (source or "").upper()

        if src == "POPUP":
            if self._capture_object_region("wnd[1]", out_path, pad=20):
                return out_path

        if src == "STATUSBAR":
            if self._capture_object_region("wnd[0]/sbar", out_path, pad=8):
                return out_path

        if src == "UNMAPPED_PARAM" and field_id:
            crop = self._capture_field_crop_from_hardcopy(field_id, out_path)
            if crop:
                return crop

        return self._hardcopy_wnd0(out_path)

    def _capture_success_evidence(self, out_path: str) -> str:
        return self._hardcopy_wnd0(out_path) if out_path else ""

    def apply_parameters_dict(self, tcode: str, params: dict[str, str]) -> tuple[dict[str, str], str]:
        """
        Tenta preencher na tela ATUAL. Se o campo estiver em outra tela, guarda para a próxima!
        """
        self._ensure_session()
        if not params:
            return {}, ""

        tcode_maps = self.field_map.get((tcode or "").upper(), {})
        full_mapping = {}
        for screen_key, fields in tcode_maps.items():
            if isinstance(fields, dict):
                full_mapping.update(fields)

        mapping_norm = {self._norm_key(k): v for k, v in full_mapping.items()}
        
        remaining_params = {}

        for key, value in params.items():
            if value is None or str(value).strip() == "":
                continue

            sap_id = full_mapping.get(key) or mapping_norm.get(self._norm_key(key))
            
            if sap_id:
                try:
                    obj = self.session.findById(sap_id)
                    
                    val_str = str(value).strip()
                    if val_str.endswith(".0"):
                        val_str = val_str[:-2]

                    obj_type = ""
                    try:
                        obj_type = str(getattr(obj, "Type", ""))
                    except Exception:
                        pass

                    if obj_type == "GuiComboBox":
                        obj.key = val_str
                    elif obj_type in ("GuiCheckBox", "GuiRadioButton"):
                        obj.selected = (val_str.upper() in["X", "1", "TRUE", "SIM", "S", "Y", "YES"])
                    else:
                        try:
                            obj.text = val_str
                        except Exception:
                            obj.key = val_str
                except Exception:
                    remaining_params[key] = value
            else:
                remaining_params[key] = value

        return remaining_params, ""

    # EXECUÇÃO
    def open_tcode(self, tcode: str) -> None:
        self._ensure_session()
        if self._popup_exists():
            self._dismiss_popup()

        self.session.findById("wnd[0]/tbar[0]/okcd").text = f"/n{tcode.upper()}"
        self.session.findById("wnd[0]").sendVKey(0)
        time.sleep(0.6)

    def execute_default(self) -> None:
        self.session.findById("wnd[0]").sendVKey(0)
        time.sleep(0.7)

    def run_tcode(
        self,
        tcode: str,
        parameters: dict,
        explanation: str,
        evidence_path: str = "",
    ) -> SapResult:
        try:
            self._ensure_session()
            self.open_tcode(tcode)
            params_to_fill = {k: v for k, v in (parameters or {}).items() if v is not None and str(v).strip() != ""}
            popup_msgs =[]
            max_telas = 10
            tela_atual = 0

            # LOOP INTELIGENTE: Navega pelas telas enquanto houver campos a preencher
            while params_to_fill and tela_atual < max_telas:
                tela_atual += 1
                
                params_to_fill, error_msg = self.apply_parameters_dict(tcode, params_to_fill)
                
                if error_msg:
                    ev = self._capture_error_evidence(evidence_path, "UNMAPPED_PARAM")
                    return SapResult("FAIL", "UNMAPPED_PARAM", error_msg, ev)

                tentativas_popup = 0
                while self._popup_exists() and tentativas_popup < 5:
                    txt = self._popup_text()
                    if txt:
                        popup_msgs.append(txt)
                    self._dismiss_popup()
                    tentativas_popup += 1

                sb_type = self._statusbar_type()
                if sb_type in {"E", "A", "X"}:
                    sb = self._statusbar_text()
                    ev = self._capture_error_evidence(evidence_path, "STATUSBAR")
                    return SapResult("FAIL", "STATUSBAR", sb or "Erro SAP", ev)

                if params_to_fill:
                    self.session.findById("wnd[0]").sendVKey(0)
                    time.sleep(0.7)

            if params_to_fill:
                campos_faltantes = list(params_to_fill.keys())
                msg = f"Campos obrigatórios ausentes ou não encontrados no SAP: {campos_faltantes}"
                ev = self._capture_error_evidence(evidence_path, "UNMAPPED_PARAM")
                return SapResult("FAIL", "UNMAPPED_PARAM", msg, ev)

            self.execute_default()
            time.sleep(1.5)
            
            tentativas_popup = 0
            while self._popup_exists() and tentativas_popup < 5:
                txt = self._popup_text()
                if txt:
                    popup_msgs.append(txt)
                self._dismiss_popup()
                tentativas_popup += 1
                time.sleep(0.5)

            tentativas_limpeza = 0
            while self._statusbar_type() in {"W", "I", "S"} and tentativas_limpeza < 3:
                try:
                    self.session.findById("wnd[0]").sendVKey(0) 
                    time.sleep(0.7)
                except Exception:
                    pass
                tentativas_limpeza += 1

            salvou = False
            botoes_salvar =[
                "wnd[0]/tbar[0]/btn[11]",
                "wnd[0]/tbar[1]/btn[11]"
            ]

            for tentativa in range(3): 
                try:
                    clicou = False
                    for btn in botoes_salvar:
                        try:
                            self.session.findById(btn).press()
                            clicou = True
                            break
                        except Exception:
                            continue
                    
                    if not clicou:
                        self.session.findById("wnd[0]").sendVKey(11)
                    
                    time.sleep(1.5)

                    teve_popup_no_save = False
                    while self._popup_exists():
                        txt = self._popup_text()
                        if txt:
                            popup_msgs.append(txt)
                        self._dismiss_popup()
                        teve_popup_no_save = True
                        time.sleep(0.5)

                    if teve_popup_no_save:
                        continue

                    salvou = True
                    break

                except Exception:
                    time.sleep(1.0)

            time.sleep(1.0)
            sb = self._statusbar_text()
            sb_type = self._statusbar_type()

            if sb_type in {"E", "A", "X"}:
                ev = self._capture_error_evidence(evidence_path, "STATUSBAR")
                return SapResult("FAIL", "STATUSBAR", sb or "Erro SAP ao salvar", ev)

            ev = self._capture_success_evidence(evidence_path)
            msg = sb or "Executado com sucesso"
            if popup_msgs:
                msg += f" | POPUP: {' || '.join(popup_msgs)}"

            return SapResult("PASS", "OK", msg, ev)

        except Exception as e:
            dump = dump_screen(self.session) if self.session else ""
            ev = self._capture_error_evidence(evidence_path, "STATUSBAR")
            return SapResult("FAIL", "EXCEPTION", f"{e} | DUMP: {dump}", ev)