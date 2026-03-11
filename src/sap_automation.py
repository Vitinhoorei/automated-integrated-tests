from __future__ import annotations

import win32com.client
import yaml
import time
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Optional
from datetime import datetime, timedelta
from sap_screen_dump import dump_screen

@dataclass
class SapResult:
    status: str
    source: str
    message: str
    evidence_path: str = ""


class SapAutomation:
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

    def _normalize_mode(self, mode: str) -> str:
        mode_n = (mode or "").strip().lower()
        if mode_n in {"simulado", "simulada", "simulate", "simulacao", "simulação"}:
            return "simulado"
        return "real"

    def _is_real_mode(self, mode: str) -> bool:
        return self._normalize_mode(mode) == "real"

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

            texts = []
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
        try:
            wnd1 = self.session.findById("wnd[1]")
        except Exception:
            return

        botoes_confirmacao = ["tbar[0]/btn[0]", "usr/btnBUTTON_1", "tbar[0]/btn[11]"]
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
            left, top, width, height = (
                int(getattr(obj, "ScreenLeft", 0)),
                int(getattr(obj, "ScreenTop", 0)),
                int(getattr(obj, "Width", 0)),
                int(getattr(obj, "Height", 0)),
            )
            if width <= 0 or height <= 0:
                return False

            bbox = (max(0, left - pad), max(0, top - pad), left + width + pad, top + height + pad)
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
                x1, y1 = int(((fx - wx) / ww) * iw) - pad, int(((fy - wy) / wh) * ih) - pad
                x2, y2 = int(((fx + fw - wx) / ww) * iw) + pad, int(((fy + fh - wy) / wh) * ih) + pad
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

    def _pending_iw31_operation_fields(self, params: dict[str, str]) -> dict[str, str]:
        op_keys = {
            "texto operação",
            "texto operacao",
            "trabalho",
            "nº colaboradores",
            "no colaboradores",
            "numero colaboradores",
        }

        out = {}
        for k, v in (params or {}).items():
            if v is None or str(v).strip() == "":
                continue
            if self._norm_key(k) in op_keys:
                out[k] = v
        return out

    def _save_current_document(self) -> bool:
        botoes_salvar = ["wnd[0]/tbar[0]/btn[11]", "wnd[0]/tbar[1]/btn[11]"]

        for btn in botoes_salvar:
            try:
                self.session.findById(btn).press()
                time.sleep(1.2)
                return True
            except Exception:
                continue

        try:
            self.session.findById("wnd[0]").sendVKey(11)
            time.sleep(1.2)
            return True
        except Exception:
            return False

    def _get_param_value(self, params: dict[str, str], *aliases: str) -> str:
        if not params:
            return ""
        aliases_norm = {self._norm_key(a) for a in aliases}
        for key, value in params.items():
            if self._norm_key(key) in aliases_norm and value is not None:
                return str(value).strip()
        return ""

    def _set_text_if_exists(self, obj_id: str, value: str) -> bool:
        if not obj_id:
            return False
        try:
            obj = self.session.findById(obj_id)
            try:
                obj.text = value
            except Exception:
                obj.key = value
            return True
        except Exception:
            return False

    def _set_checkbox_if_exists(self, obj_id: str, checked: bool) -> bool:
        if not obj_id:
            return False
        try:
            obj = self.session.findById(obj_id)
            obj.selected = checked
            return True
        except Exception:
            return False

    def _table_exists(self, obj_id: str) -> bool:
        try:
            self.session.findById(obj_id)
            return True
        except Exception:
            return False

    def _safe_press_save(self, mode: str = "real") -> None:
        if self._is_real_mode(mode):
            self._save_current_document()
            time.sleep(0.8)

            while self._popup_exists():
                self._dismiss_popup()
                time.sleep(0.3)
        else:
            try:
                self.session.findById("wnd[0]").sendVKey(0)
                time.sleep(0.5)
            except Exception:
                pass

    def _run_iw41_flow(self, parameters: dict, evidence_path: str = "", mode: str = "real") -> SapResult:
        """
        Fluxo especial IW41:
        - entra na transação
        - Enter / Enter
        - se houver tabela com uma ou mais linhas, processa todas
        - se entrar direto no detalhe, processa uma vez
        """
        try:
            is_real_mode = self._is_real_mode(mode)

            mapping = self.field_map.get("IW41", {}).get("SAPLCORU|0100", {})
            mapping_norm = {self._norm_key(k): v for k, v in mapping.items()}

            ordem_id = mapping.get("Ordem") or mapping_norm.get(self._norm_key("Ordem"))
            pernr_id = mapping.get("Nº pessoal") or mapping_norm.get(self._norm_key("Nº pessoal"))
            conf_final_id = mapping.get("Conf.final") or mapping_norm.get(self._norm_key("Conf.final"))
            baixa_res_id = mapping.get("Dar baixa res.") or mapping_norm.get(self._norm_key("Dar baixa res."))
            isdd_id = mapping.get("Início trabalho data") or mapping_norm.get(self._norm_key("Início trabalho data"))
            isdz_id = mapping.get("Início trabalho hora") or mapping_norm.get(self._norm_key("Início trabalho hora"))
            iedd_id = mapping.get("Fim trabalho data") or mapping_norm.get(self._norm_key("Fim trabalho data"))
            iedz_id = mapping.get("Fim trabalho hora") or mapping_norm.get(self._norm_key("Fim trabalho hora"))
            table_id = mapping.get("Tabela confirmações") or mapping_norm.get(self._norm_key("Tabela confirmações"))

            ordem = self._get_param_value(parameters, "Ordem")
            pernr = self._get_param_value(parameters, "Nº pessoal", "N* pessoal", "N pessoal", "Numero pessoal")

            if ordem and ordem_id:
                self._set_text_if_exists(ordem_id, ordem)

            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.8)
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.8)

            ontem = (datetime.now() - timedelta(days=1)).strftime("%d.%m.%Y")
            hora_inicio_dt = datetime.strptime("13:00", "%H:%M")
            hora_fim_dt = hora_inicio_dt + timedelta(hours=2)
            hora_inicio = hora_inicio_dt.strftime("%H:%M")
            hora_fim = hora_fim_dt.strftime("%H:%M")
            processed = 0

            if table_id and self._table_exists(table_id):
                table = self.session.findById(table_id)

                try:
                    row_count = int(getattr(table, "RowCount", 0))
                except Exception:
                    row_count = 0

                if row_count <= 0:
                    row_count = 10 

                for row_idx in range(row_count):
                    try:
                        table.getAbsoluteRow(row_idx).selected = True
                    except Exception:
                        break

                    try:
                        icon_id = f"wnd[0]/usr/tblSAPLCORUTC_3100/txtCORUF-UPD_ICON[0,{row_idx}]"
                        self.session.findById(icon_id).setFocus()
                        self.session.findById(icon_id).caretPosition = 0
                        self.session.findById("wnd[0]").sendVKey(0)
                        time.sleep(0.5)
                    except Exception:
                        continue

                    try:
                        chk_id = f"wnd[0]/usr/tblSAPLCORUTC_3100/chkCORUF-FLG_SPL[3,{row_idx}]"
                        self.session.findById(chk_id).setFocus()
                        self.session.findById("wnd[0]").sendVKey(2)
                        time.sleep(0.7)
                    except Exception:
                        continue

                    self._set_checkbox_if_exists(conf_final_id, False)
                    self._set_checkbox_if_exists(baixa_res_id, False)

                    if pernr and pernr_id:
                        self._set_text_if_exists(pernr_id, pernr)

                    self._set_text_if_exists(isdd_id, ontem)
                    self._set_text_if_exists(isdz_id, hora_inicio)
                    self._set_text_if_exists(iedd_id, ontem)
                    self._set_text_if_exists(iedz_id, hora_fim)

                    self.session.findById("wnd[0]").sendVKey(0)
                    time.sleep(0.7)

                    self._safe_press_save(mode=mode)

                    sb_type = self._statusbar_type()
                    sb = self._statusbar_text()
                    if is_real_mode and sb_type in {"E", "A", "X"}:
                        ev = self._capture_error_evidence(evidence_path, "STATUSBAR")
                        return SapResult("FAIL", "STATUSBAR", sb or f"Erro SAP ao gravar linha {row_idx + 1} no IW41.", ev)

                    try:
                        self.session.findById("wnd[0]").sendVKey(0)
                        time.sleep(0.6)
                    except Exception:
                        pass

                    processed += 1

                ev = self._capture_success_evidence(evidence_path)
                modo_txt = "REAL" if is_real_mode else "SIMULADO"
                return SapResult("PASS", "OK", f"IW41 executado com sucesso em {processed} linha(s). ({modo_txt})", ev)

            self._set_checkbox_if_exists(conf_final_id, False)
            self._set_checkbox_if_exists(baixa_res_id, False)

            if pernr and pernr_id:
                self._set_text_if_exists(pernr_id, pernr)

            self._set_text_if_exists(isdd_id, ontem)
            self._set_text_if_exists(isdz_id, hora_inicio)
            self._set_text_if_exists(iedd_id, ontem)
            self._set_text_if_exists(iedz_id, hora_fim)

            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.7)

            self._safe_press_save(mode=mode)

            sb_type = self._statusbar_type()
            sb = self._statusbar_text()
            if is_real_mode and sb_type in {"E", "A", "X"}:
                ev = self._capture_error_evidence(evidence_path, "STATUSBAR")
                return SapResult("FAIL", "STATUSBAR", sb or "Erro SAP no IW41.", ev)

            ev = self._capture_success_evidence(evidence_path)
            modo_txt = "REAL" if is_real_mode else "SIMULADO"
            return SapResult("PASS", "OK", (sb or f"IW41 executado com sucesso. ({modo_txt})"), ev)

        except Exception as e:
            dump = dump_screen(self.session) if self.session else ""
            ev = self._capture_error_evidence(evidence_path, "STATUSBAR")
            return SapResult("FAIL", "EXCEPTION", f"{e} | DUMP: {dump}", ev)

    def _run_ip41_flow(self, parameters: dict, explanation: str, evidence_path: str = "") -> SapResult:
        """
        Fluxo especial da IP41:
        Preenche a tela inicial -> Abas -> Cria Lista de Tarefas -> Tabela -> Simulado ou Real
        """
        try:
            is_simulated = "simulado" in (explanation or "").lower()
            mapping = {}
            for k, v in self.field_map.get("IP41", {}).items():
                if isinstance(v, dict):
                    mapping.update(v)
            mapping_norm = {self._norm_key(k): v for k, v in mapping.items()}

            def get_val(key):
                for pk, pv in parameters.items():
                    if self._norm_key(pk) == self._norm_key(key):
                        return pv
                return ""

            def get_id(key):
                return mapping.get(key) or mapping_norm.get(self._norm_key(key))

            ctg = get_val("Ctg.plano manutenção") or get_val("Ctg.plano manut.")
            if ctg:
                self.session.findById(get_id("Ctg.plano manutenção")).key = str(ctg).strip()
            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.8)

            self.session.findById(get_id("Texto do plano de manutenção")).text = get_val("Texto do plano de manutenção")
            self.session.findById(get_id("Ciclo")).text = get_val("Ciclo")
            self.session.findById(get_id("Unidade do ciclo")).text = get_val("Unidade do ciclo")
            
            loc = get_val("Local de instalação")
            if loc:
                self.session.findById(get_id("Local de instalação")).text = loc
            
            equip = get_val("Equipamento")
            if equip:
                self.session.findById(get_id("Equipamento")).text = equip
                
            self.session.findById(get_id("Tipo de ordem")).text = get_val("Tipo de ordem")
            self.session.findById(get_id("Tipo de atividade de manutenção")).text = get_val("Tipo de atividade de manutenção")
            self.session.findById("wnd[0]").sendVKey(0)
            
            time.sleep(0.8)
            
            prio = get_val("Prioridade")
            if prio:
                self.session.findById(get_id("Prioridade")).key = str(prio).strip()

            self.session.findById("wnd[0]/usr/subSUBSCREEN_MPLAN:SAPLIWP3:8001/tabsTABSTRIP_HEAD/tabpT\\03").select()
            time.sleep(0.5)
            self.session.findById(get_id("Tipo de programação")).key = get_val("Tipo de programação")
            self.session.findById("wnd[0]/usr/subSUBSCREEN_MPLAN:SAPLIWP3:8001/tabsTABSTRIP_HEAD/tabpT\\01").select()
            time.sleep(0.5)
            self.session.findById(get_id("Criar lista de tarefas")).press()
            time.sleep(0.8)
            
            if self._popup_exists():
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press() 
                time.sleep(0.8)

            self.session.findById(get_id("Utilização")).text = get_val("Utilização")
            self.session.findById(get_id("Grupo de planejamento")).text = get_val("Grupo de planejamento")
            self.session.findById(get_id("Status do plano")).text = str(get_val("Status do plano")).strip()
            self.session.findById("wnd[0]/tbar[1]/btn[16]").press()
            time.sleep(0.8)
            
            self.session.findById(get_id("Descrição da operação")).text = get_val("Descrição da operação")
            self.session.findById(get_id("Trabalho")).text = get_val("Trabalho")
            self.session.findById(get_id("Unidade trabalho")).text = get_val("Unidade trabalho")
            self.session.findById(get_id("Duração")).text = get_val("Duração")
            self.session.findById(get_id("Unidade duração")).text = get_val("Unidade duração")
            self.session.findById("wnd[0]").sendVKey(0)
            
            time.sleep(1.0)
            ev = self._capture_success_evidence(evidence_path)

            if is_simulated:
                msg = "Fluxo IP41 executado com sucesso. (SIMULADO - Operação cancelada no final)"
                try:
                    self.session.findById("wnd[0]/tbar[0]/btn[15]").press() 
                    time.sleep(0.5)
                    if self._popup_exists():
                        self.session.findById("wnd[1]/usr/btnSPOP-OPTION2").press() 
                    time.sleep(0.5)
                    self.session.findById("wnd[0]/tbar[0]/btn[15]").press()  
                    time.sleep(0.5)
                    if self._popup_exists():
                        self.session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
                except Exception:
                    pass
                return SapResult("PASS", "OK", msg, ev)
                
            else:
                msg = "Fluxo IP41 executado e salvo com sucesso."
                self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
                time.sleep(1.5)
                
                sb = self._statusbar_text()
                sb_type = self._statusbar_type()

                if sb_type in {"E", "A", "X"}:
                    ev = self._capture_error_evidence(evidence_path, "STATUSBAR")
                    return SapResult("FAIL", "STATUSBAR", sb or "Erro SAP ao salvar IP41", ev)
                
                return SapResult("PASS", "OK", sb or msg, ev)

        except Exception as e:
            ev = self._capture_error_evidence(evidence_path, "STATUSBAR")
            return SapResult("FAIL", "EXCEPTION", f"Erro no fluxo IP41: {e}", ev)
        
    def apply_parameters_dict(self, tcode: str, params: dict[str, str]) -> tuple[dict[str, str], str, str]:
        """
        Retorna (remaining_params, error_msg, action_taken).
        action_taken pode ser: "TEXT", "BUTTON", "TAB" ou "NONE".
        Isso garante que o robô faça apenas UMA COISA por vez na tela.
        """
        self._ensure_session()
        if not params:
            return {}, "", "NONE"

        tcode_maps = self.field_map.get((tcode or "").upper(), {})
        full_mapping = {}
        for screen_key, fields in tcode_maps.items():
            if isinstance(fields, dict):
                full_mapping.update(fields)

        mapping_norm = {self._norm_key(k): v for k, v in full_mapping.items()}
        remaining_params = {}
        campos_normais = {}
        abas = {}
        botoes = {}

        for key, value in params.items():
            if value is None or str(value).strip() == "":
                continue
            sap_id = full_mapping.get(key) or mapping_norm.get(self._norm_key(key))
            if sap_id:
                try:
                    obj = self.session.findById(sap_id)
                    obj_type = str(getattr(obj, "Type", ""))
                    if obj_type == "GuiTab":
                        abas[key] = (obj, value)
                    elif obj_type == "GuiButton":
                        botoes[key] = (obj, value)
                    else:
                        campos_normais[key] = (obj, obj_type, value)
                except Exception:
                    remaining_params[key] = value
            else:
                remaining_params[key] = value

        action_taken = "NONE"

        if campos_normais:
            for key, (obj, obj_type, value) in campos_normais.items():
                val_str = str(value).strip()
                if val_str.endswith(".0"):
                    val_str = val_str[:-2]
                try:
                    if obj_type == "GuiComboBox":
                        obj.key = val_str
                    elif obj_type in ("GuiCheckBox", "GuiRadioButton"):
                        obj.selected = (val_str.upper() in ["X", "1", "TRUE", "SIM", "S", "Y", "YES"])
                    else:
                        try:
                            obj.text = val_str
                        except Exception:
                            obj.key = val_str
                    action_taken = "TEXT"
                except Exception:
                    remaining_params[key] = value

            for k, v in abas.items():
                remaining_params[k] = v[1]
            for k, v in botoes.items():
                remaining_params[k] = v[1]
            return remaining_params, "", action_taken

        if botoes:
            clicked_one = False
            for key, (obj, value) in botoes.items():
                if not clicked_one:
                    val_str = str(value).strip().upper()
                    if val_str in ["X", "1", "TRUE", "SIM", "S", "Y", "YES"]:
                        try:
                            obj.press()
                            action_taken = "BUTTON"
                            clicked_one = True
                            time.sleep(0.9)
                        except Exception:
                            remaining_params[key] = value
                    else:
                        remaining_params[key] = value
                else:
                    remaining_params[key] = value

            for k, v in abas.items():
                remaining_params[k] = v[1]
            return remaining_params, "", action_taken

        if abas:
            clicked_one = False
            for key, (obj, value) in abas.items():
                if not clicked_one:
                    try:
                        obj.select()
                        action_taken = "TAB"
                        clicked_one = True
                        time.sleep(0.5)
                    except Exception:
                        remaining_params[key] = value
                else:
                    remaining_params[key] = value
            return remaining_params, "", action_taken

        return remaining_params, "", "NONE"

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
        mode: str = "real",
    ) -> SapResult:
        try:
            self._ensure_session()
            self.open_tcode(tcode)

            exec_mode = self._normalize_mode(mode)
            is_real_mode = self._is_real_mode(exec_mode)

            if tcode.upper() == "IW41":
                return self._run_iw41_flow(parameters, evidence_path, mode=exec_mode)
            
            if tcode.upper() == "IP41":
                return self._run_ip41_flow(parameters, explanation, evidence_path)

            params_to_fill = {k: v for k, v in (parameters or {}).items() if v is not None and str(v).strip() != ""}
            popup_msgs = []
            max_telas = 10
            tela_atual = 0

            has_iw31_operation_flow = tcode.upper() == "IW31" and bool(self._pending_iw31_operation_fields(params_to_fill))
            iw31_operations_enter_done = False
            iw31_operations_save_done = False
            is_iw32_print_flow = tcode.upper() == "IW32" and "imprimir" in (explanation or "").lower()

            while params_to_fill and tela_atual < max_telas:
                tela_atual += 1
                # print(f"[DEBUG] Tela {tela_atual} - Parâmetros sobrando: {list(params_to_fill.keys())}")
                params_to_fill, error_msg, action_taken = self.apply_parameters_dict(tcode, params_to_fill)

                if has_iw31_operation_flow and tcode.upper() == "IW31":
                    pending_ops = self._pending_iw31_operation_fields(params_to_fill)

                    if not pending_ops:
                        if not iw31_operations_enter_done:
                            self.session.findById("wnd[0]").sendVKey(0)
                            time.sleep(0.8)
                            iw31_operations_enter_done = True
                            continue

                        if iw31_operations_enter_done and not iw31_operations_save_done:
                            if is_real_mode:
                                try:
                                    self.session.findById("wnd[0]/tbar[1]/btn[25]").press()
                                    time.sleep(1.0)
                                except Exception:
                                    pass

                            saved = True
                            if is_real_mode:
                                saved = self._save_current_document()

                            if saved:
                                iw31_operations_save_done = True
                                time.sleep(1.0)

                                while self._popup_exists():
                                    txt = self._popup_text()
                                    if txt:
                                        popup_msgs.append(txt)
                                    self._dismiss_popup()
                                    time.sleep(0.4)

                                sb = self._statusbar_text()
                                sb_type = self._statusbar_type()

                                if is_real_mode and sb_type in {"E", "A", "X"}:
                                    ev = self._capture_error_evidence(evidence_path, "STATUSBAR")
                                    return SapResult("FAIL", "STATUSBAR", sb or "Erro SAP ao salvar", ev)

                                ev = self._capture_success_evidence(evidence_path)
                                modo_txt = "REAL" if is_real_mode else "SIMULADO"
                                msg = sb or f"Executado com sucesso ({modo_txt})"
                                if popup_msgs:
                                    msg += f" | POPUP: {' || '.join(popup_msgs)}"
                                return SapResult("PASS", "OK", msg, ev)

                if error_msg:
                    ev = self._capture_error_evidence(evidence_path, "UNMAPPED_PARAM")
                    return SapResult("FAIL", "UNMAPPED_PARAM", error_msg, ev)

                tentativas_popup = 0
                popups_dismissed = False
                while self._popup_exists() and tentativas_popup < 5:
                    txt = self._popup_text()
                    if txt:
                        popup_msgs.append(txt)
                    self._dismiss_popup()
                    tentativas_popup += 1
                    popups_dismissed = True

                sb_type = self._statusbar_type()
                if sb_type in {"E", "A", "X"}:
                    sb = self._statusbar_text()
                    ev = self._capture_error_evidence(evidence_path, "STATUSBAR")
                    return SapResult("FAIL", "STATUSBAR", sb or "Erro SAP", ev)

                if params_to_fill:
                    if not popups_dismissed:
                        if action_taken == "TEXT":
                            self.session.findById("wnd[0]").sendVKey(0)
                            time.sleep(0.7)
                        elif action_taken == "NONE":
                            if (
                                has_iw31_operation_flow
                                and tcode.upper() == "IW31"
                                and iw31_operations_enter_done
                            ):
                                time.sleep(0.3)
                            elif str(self.session.Info.Program).strip() == "SAPLIQS0" and tcode.upper() == "IW31":
                                self.session.findById("wnd[0]").sendVKey(3)
                                time.sleep(0.7)
                            else:
                                self.session.findById("wnd[0]").sendVKey(0)
                                time.sleep(0.7)

            if params_to_fill:
                campos_faltantes = list(params_to_fill.keys())
                msg = f"Campos obrigatórios ausentes ou não encontrados no SAP: {campos_faltantes}"
                ev = self._capture_error_evidence(evidence_path, "UNMAPPED_PARAM")
                return SapResult("FAIL", "UNMAPPED_PARAM", msg, ev)

            self.execute_default()
            time.sleep(1.5)

            if is_iw32_print_flow and tcode.upper() == "IW32":
                try:
                    self.session.findById("wnd[0]/tbar[0]/btn[86]").press()
                    time.sleep(1.5)

                    self.session.findById("wnd[1]/usr/tblSAPLIPRTTC_WORKPAPERS").getAbsoluteRow(8).selected = True
                    self.session.findById("wnd[1]/tbar[0]/btn[16]").press()
                    time.sleep(3.0)
                    ev = self._capture_success_evidence(evidence_path)
                    msg = "Visualização de impressão gerada com sucesso na tela."

                    try:
                        self.session.findById("wnd[0]/tbar[0]/btn[12]").press()
                        time.sleep(0.8)

                        if self._popup_exists():
                            try:
                                self.session.findById("wnd[1]").close()
                            except Exception:
                                self._dismiss_popup()
                            time.sleep(0.5)

                        self.session.findById("wnd[0]/tbar[0]/btn[12]").press()
                        time.sleep(0.8)

                        if self._popup_exists():
                            self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                    except Exception:
                        pass

                    modo_txt = "REAL" if is_real_mode else "SIMULADO"
                    return SapResult("PASS", "OK", f"{msg} ({modo_txt})", ev)

                except Exception as e:
                    ev = self._capture_error_evidence(evidence_path, "STATUSBAR")
                    return SapResult("FAIL", "EXCEPTION", f"Falha na impressão IW32: {e}", ev)

            while self._popup_exists():
                txt = self._popup_text()
                if txt:
                    popup_msgs.append(txt)
                self._dismiss_popup()
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

            if is_real_mode:
                botoes_salvar = ["wnd[0]/tbar[0]/btn[11]", "wnd[0]/tbar[1]/btn[11]"]
                for tentativa in range(3):
                    try:
                        if tcode.upper() == "IW31":
                            try:
                                self.session.findById("wnd[0]/tbar[1]/btn[25]").press()
                                time.sleep(1.0)
                            except Exception:
                                pass

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
                        while self._popup_exists():
                            txt = self._popup_text()
                            if txt:
                                popup_msgs.append(txt)
                            self._dismiss_popup()
                            time.sleep(0.5)

                        salvou = True
                        break
                    except Exception:
                        time.sleep(1.0)
            else:
                salvou = True

            time.sleep(1.0)
            sb = self._statusbar_text()
            sb_type = self._statusbar_type()

            if is_real_mode and sb_type in {"E", "A", "X"}:
                ev = self._capture_error_evidence(evidence_path, "STATUSBAR")
                return SapResult("FAIL", "STATUSBAR", sb or "Erro SAP ao salvar", ev)

            ev = self._capture_success_evidence(evidence_path)
            modo_txt = "REAL" if is_real_mode else "SIMULADO"
            msg = sb or f"Executado com sucesso ({modo_txt})"
            if popup_msgs:
                msg += f" | POPUP: {' || '.join(popup_msgs)}"
            return SapResult("PASS", "OK", msg, ev)

        except Exception as e:
            dump = dump_screen(self.session) if self.session else ""
            ev = self._capture_error_evidence(evidence_path, "STATUSBAR")
            return SapResult("FAIL", "EXCEPTION", f"{e} | DUMP: {dump}", ev)