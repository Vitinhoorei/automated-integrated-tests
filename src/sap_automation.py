from __future__ import annotations

import time
import unicodedata
import re
import win32com.client
import yaml
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional
from sap_screen_dump import dump_screen
from param_enricher import enrich_params

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
        
    def go_to_initial_screen(self) -> None:
        """Força a volta para a tela inicial limpando inclusive popups de saída."""
        self._ensure_session()
        try:
            for _ in range(3):
                while self._popup_exists():
                    self._dismiss_popup()
                    time.sleep(0.3)
                
                self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                self.session.findById("wnd[0]").sendVKey(0)
                time.sleep(0.5)
                
                if not self._popup_exists():
                    break
                
        except Exception:
            pass

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

        if mode_n in {"executar", "exec", "real"}:
            return "real"

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

    def _dismiss_popup(self, button_type: str = "YES") -> None:
        try:
            wnd1 = self.session.findById("wnd[1]")
        except Exception:
            return

        if button_type == "NO":
            candidates = ["usr/btnSPOP-OPTION2", "usr/btnBUTTON_2"]
        else:
            candidates = ["usr/btnSPOP-OPTION1", "usr/btnBUTTON_1", "tbar[0]/btn[0]", "tbar[0]/btn[11]"]
        
        for btn in candidates:
            try:
                wnd1.findById(btn).press()
                time.sleep(0.4)
                return
            except Exception:
                continue
        wnd1.sendVKey(0)

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
            "texto operação 2",
            "texto operacao 2",
            "trabalho 2",
            "nº colaboradores 2",
            "no colaboradores 2",
            "numero colaboradores 2",
            "texto operação 3",
            "texto operacao 3",
            "trabalho 3",
            "nº colaboradores 3",
            "no colaboradores 3",
            "numero colaboradores 3",
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

    def _handle_iw21_z4_popup(self, parameters: dict[str, str]) -> bool:
        """
        Se for nota Z4 e aparecer popup na IW21, tenta clicar em Avançar.
        """
        try:
            if self._screen_key() != "SAPLIQS0|0100":
                return False

            tipo_nota = self._get_param_value(parameters, "Tipo de nota").upper().strip()
            if tipo_nota != "Z4":
                return False

            if not self._popup_exists():
                return False

            iw21_maps = self.field_map.get("IW21", {})
            popup_map = iw21_maps.get("POPUP_Z4", {}) if isinstance(iw21_maps, dict) else {}
            avancar_id = ""
            if isinstance(popup_map, dict):
                avancar_id = popup_map.get("Avançar", "") or popup_map.get("Avancar", "")

            candidates = []
            if avancar_id:
                candidates.append(avancar_id)

            candidates.extend([
                "wnd[1]/tbar[0]/btn[0]",
                "wnd[1]/usr/btnBUTTON_1",
                "wnd[1]/usr/btnSPOP-OPTION1",
            ])

            for obj_id in candidates:
                try:
                    self.session.findById(obj_id).press()
                    time.sleep(0.8)
                    return True
                except Exception:
                    continue

            try:
                self.session.findById("wnd[1]").sendVKey(0)
                time.sleep(0.8)
                return True
            except Exception:
                return False

        except Exception:
            return False

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
            hora_inicio_dt = datetime.strptime("16:00", "%H:%M")
            hora_fim_dt = hora_inicio_dt + timedelta(hours=2)
            hora_inicio = hora_inicio_dt.strftime("%H:%M")
            hora_fim = hora_fim_dt.strftime("%H:%M")
            processed = 0

            if table_id and self._table_exists(table_id):
                row_idx = 0

                while True:
                    if not self._table_exists(table_id):
                        break

                    try:
                        table = self.session.findById(table_id)
                        row_count = int(getattr(table, "RowCount", 0))
                    except Exception:
                        break

                    if row_count <= 0 or row_idx >= row_count:
                        break

                    try:
                        table.getAbsoluteRow(row_idx).selected = True
                        time.sleep(0.2)
                    except Exception:
                        break

                    entrou_linha = False
                    for icon_candidate in [
                        f"wnd[0]/usr/tblSAPLCORUTC_3100/txtCORUF-UPD_ICON[0,{row_idx}]",
                        f"wnd[0]/usr/tblSAPLCORUTC_3100/chkCORUF-FLG_SPL[3,{row_idx}]",
                    ]:
                        try:
                            obj = self.session.findById(icon_candidate)
                            obj.setFocus()
                            try:
                                obj.caretPosition = 0
                            except Exception:
                                pass
                            self.session.findById("wnd[0]").sendVKey(0)
                            time.sleep(0.5)
                            entrou_linha = True
                            break
                        except Exception:
                            continue

                    if not entrou_linha:
                        row_idx += 1
                        continue

                    abriu_detalhe = False
                    try:
                        chk_id = f"wnd[0]/usr/tblSAPLCORUTC_3100/chkCORUF-FLG_SPL[3,{row_idx}]"
                        self.session.findById(chk_id).setFocus()
                        self.session.findById("wnd[0]").sendVKey(2)
                        time.sleep(0.7)
                        abriu_detalhe = True
                    except Exception:
                        pass

                    if not abriu_detalhe:
                        row_idx += 1
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

                    processed += 1

                    voltou = False
                    for _ in range(3):
                        try:
                            self.session.findById("wnd[0]").sendVKey(0)
                            time.sleep(0.6)
                        except Exception:
                            pass

                        if self._table_exists(table_id):
                            voltou = True
                            break

                        try:
                            self.session.findById("wnd[0]").sendVKey(12)
                            time.sleep(0.6)
                        except Exception:
                            pass

                        if self._table_exists(table_id):
                            voltou = True
                            break

                    if not voltou and not self._table_exists(table_id):
                        break

                    row_idx += 1

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

    def _run_ip41_ip42_flow(self, tcode: str, parameters: dict, explanation: str, evidence_path: str = "", mode: str = "real") -> SapResult:
        """
        Fluxo especial Unificado da IP41 e IP42:
        Preenche a tela inicial -> Abas -> Cria Lista de Tarefas -> Tabela -> PactsMt (Apenas IP42) -> Simulado ou Real
        """
        try:
            tcode_u = tcode.upper()
            is_real_mode = self._is_real_mode(mode)
            mapping = {}
            for k, v in self.field_map.get(tcode_u, {}).items():
                if isinstance(v, dict):
                    mapping.update(v)
            mapping_norm = {self._norm_key(k): v for k, v in mapping.items()}

            def get_val(key):
                for pk, pv in parameters.items():
                    if self._norm_key(pk) == self._norm_key(key):
                        return str(pv).strip() if pv is not None else ""
                return ""

            def get_id(key):
                sap_id = mapping.get(key) or mapping_norm.get(self._norm_key(key))
                if not sap_id:
                    raise ValueError(f"ID não mapeado no YAML da {tcode_u} para o campo: {key}")
                return sap_id

            def safe_set_text(key_name):
                val = get_val(key_name)
                if val:
                    try:
                        self.session.findById(get_id(key_name)).text = val
                    except Exception:
                        raise Exception(f"Campo '{key_name}' não encontrado na tela. O SAP travou ou a tela não carregou.")

            def safe_set_key(key_name):
                val = get_val(key_name)
                if val:
                    try:
                        self.session.findById(get_id(key_name)).key = val
                    except Exception:
                        raise Exception(f"Combo '{key_name}' não encontrado na tela. O SAP travou.")

            ctg = get_val("Ctg.plano de manutenção") or get_val("Ctg.plano manutenção") or get_val("Ctg.plano manut.")
            if ctg:
                try:
                    id_ctg = mapping.get("Ctg.plano manut.") or mapping.get("Ctg.plano de manutenção") or mapping_norm.get(self._norm_key("Ctg.plano manutenção"))
                    self.session.findById(id_ctg).key = ctg
                except Exception:
                    raise Exception("Campo 'Ctg.plano de manutenção' não encontrado na tela inicial.")

            if tcode_u == "IP42":
                estra = get_val("Estratégia") or get_val("Estrategia")
                if estra:
                    try:
                        self.session.findById(get_id("Estratégia")).text = estra
                    except Exception:
                        raise Exception("Campo 'Estratégia' não encontrado na tela inicial (IP42).")

            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.8)
            val_texto = get_val("Texto do plano de manutenção") or get_val("Texto Breve")
            if val_texto:
                try:
                    self.session.findById(get_id("Texto do plano de manutenção")).text = val_texto
                except Exception:
                    raise Exception("Campo 'Texto do plano de manutenção' não encontrado.")

            if tcode_u == "IP41":
                safe_set_text("Ciclo")
                safe_set_text("Unidade do ciclo")

            safe_set_text("Local de instalação")

            val_equip = get_val("Nº equipamento") or get_val("Equipamento")
            if val_equip:
                try:
                    self.session.findById(get_id("Nº equipamento")).text = val_equip
                except Exception:
                    raise Exception("Campo 'Nº equipamento' não encontrado.")

            safe_set_text("Tipo de ordem")
            safe_set_text("Tipo de atividade de manutenção")

            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.8)
            safe_set_key("Prioridade")

            try:
                self.session.findById("wnd[0]/usr/subSUBSCREEN_MPLAN:SAPLIWP3:8001/tabsTABSTRIP_HEAD/tabpT\\03").select()
                time.sleep(0.5)

                val_plan_sort = get_val("Tipo de programação") or get_val("Campo ordenação") or get_val("Campo ordenacao")
                if val_plan_sort:
                    try:
                        self.session.findById(get_id("Campo Ordenação")).key = val_plan_sort
                    except Exception:
                        self.session.findById(get_id("Campo Ordenação")).value = val_plan_sort

                self.session.findById("wnd[0]").sendVKey(0)
                time.sleep(0.5)
            except Exception:
                pass

            self.session.findById("wnd[0]/usr/subSUBSCREEN_MPLAN:SAPLIWP3:8001/tabsTABSTRIP_HEAD/tabpT\\01").select()
            time.sleep(0.5)

            try:
                self.session.findById(get_id("Criar lista de tarefas")).press()
            except Exception:
                raise Exception("Botão 'Criar lista de tarefas' não encontrado.")
            time.sleep(1.0)

            if self._popup_exists():
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                time.sleep(0.8)

            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1.0)

            safe_set_text("Utilização")
            safe_set_text("Grupo de planejamento")
            safe_set_text("Status do plano")

            try:
                self.session.findById("wnd[0]/tbar[1]/btn[16]").press()
            except Exception:
                raise Exception("Botão de Operações (btn[16]) não encontrado.")
            time.sleep(1.0)

            val_desc_op = get_val("Descrição da operação") or get_val("Descrição operação") or get_val("Texto Operação")
            if val_desc_op:
                try:
                    self.session.findById(get_id("Descrição da operação")).text = val_desc_op
                except Exception:
                    raise Exception("Campo 'Descrição da operação' na tabela não encontrado.")

            safe_set_text("Trabalho")
            safe_set_text("Unidade trabalho")
            safe_set_text("Duração")
            safe_set_text("Unidade duração")

            self.session.findById("wnd[0]").sendVKey(0)
            time.sleep(1.2)

            if tcode_u == "IP42":
                try:
                    self.session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_3400").getAbsoluteRow(0).selected = True
                    time.sleep(0.3)
                    self.session.findById("wnd[0]/usr/btnTEXT_DRUCKTASTE_WP").press()
                    time.sleep(1.0)
                    self.session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_3600/chkRIHSTRAT-MARK01[3,0]").selected = True
                    time.sleep(0.3)
                    self.session.findById("wnd[0]/tbar[1]/btn[32]").press()
                    time.sleep(1.0)
                except Exception as e:
                    raise Exception(f"Erro ao atribuir Pacotes de Manutenção (PactsMt) na IP42: {e}")

            ev = self._capture_success_evidence(evidence_path)

            if not is_real_mode:
                msg = f"Fluxo {tcode_u} executado com sucesso. (SIMULADO - Operação cancelada no final)"
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
                msg = f"Fluxo {tcode_u} executado e salvo com sucesso."
                self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
                time.sleep(1.5)

                sb = self._statusbar_text()
                sb_type = self._statusbar_type()

                if sb_type in {"E", "A", "X"}:
                    ev = self._capture_error_evidence(evidence_path, "STATUSBAR")
                    return SapResult("FAIL", "STATUSBAR", sb or f"Erro SAP ao salvar {tcode_u}", ev)

                return SapResult("PASS", "OK", sb or msg, ev)

        except Exception as e:
            ev = self._capture_error_evidence(evidence_path, "STATUSBAR")
            return SapResult("FAIL", "EXCEPTION", str(e), ev)
        
    # PP
    def _run_pp_order_flow(self, tcode, parameters, evidence_path, mode):
        tcode_u = tcode.upper()
        self.open_tcode(tcode_u)
        fluxo = parameters.get("_FLUXO", "PADRAO")
        time.sleep(1.0)
        self.apply_parameters_dict(tcode_u, parameters)
        
        if tcode_u == "CO02":
            for sub_key in ["CO02_DATAS", "CO02_LIBERAR", "CO02_TECO"]:
                if sub_key in self.field_map:
                    self.apply_parameters_dict(sub_key, parameters)

        self.session.findById("wnd[0]").sendVKey(0) 
        time.sleep(1.0)
        
        if tcode_u == "CO02":
            self.session.findById("wnd[0]").sendVKey(0) 
            time.sleep(0.8)

        if tcode_u == "CO01":
            self.apply_parameters_dict("CO01", parameters)
            self.session.findById("wnd[0]").sendVKey(0) 
            
        elif tcode_u == "CO02":
            if fluxo == "DATAS":
                try:
                    self.session.findById(self.field_map["CO02_DATAS"]["SAPLCOKO1|0120"]["Data fim"]).text = ""
                except: pass
                
                self.apply_parameters_dict("CO02_DATAS", parameters)
                self.session.findById("wnd[0]").sendVKey(0) 
                
                self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
                time.sleep(0.8)
                if self._popup_exists(): 
                    self._dismiss_popup("NO") 
                
                self.session.findById("wnd[0]/tbar[0]/btn[11]").press() 
                
            elif fluxo == "LIBERAR_IMPRIMIR":
                self.session.findById(self.field_map["CO02_LIBERAR"]["metadata"]["btn_liberar"]).press()
                if self._popup_exists(): 
                    self.session.findById(self.field_map["CO02_LIBERAR"]["metadata"]["btn_popup_nao"]).press()
                
                self.session.findById("wnd[0]/tbar[0]/btn[11]").press() 
                
            elif fluxo == "TECO":
                self.session.findById(self.field_map["CO02_TECO"]["metadata"]["menu_teco"]).select()
                self.session.findById("wnd[0]/tbar[0]/btn[11]").press()

        return self._finalizar_pelo_modo_universal(tcode_u, mode, evidence_path)

    def _run_pp_co11n_flow(self, parameters, evidence_path, mode):
        ops = parameters.get("_LISTA_OPERACOES", ["0010"])
        for op in ops:
            self.open_tcode("CO11N")
            self.apply_parameters_dict("CO11N", {"Ordem": parameters.get("Ordem"), "Operação": op})
            self.session.findById("wnd[0]").sendVKey(0)
            self._finalizar_pelo_modo_universal("CO11N", mode, evidence_path)
        return SapResult("PASS", "OK", "Apontamento(s) realizado(s)")

    def _run_pp_co07_flow(self, parameters, evidence_path, mode):
        self.open_tcode("CO07")
        self.apply_parameters_dict("CO07", parameters)
        self.session.findById("wnd[0]").sendVKey(0)
        self.apply_parameters_dict("CO07", parameters)
        self.session.findById("wnd[0]").sendVKey(0)
        
        if self._popup_exists():
            self.session.findById(self.field_map["CO07"]["metadata"]["btn_gerar_operacao"]).press()
            time.sleep(1.0)
            
        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
        if self._popup_exists():
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            
        self.session.findById(self.field_map["CO07"]["SAPLCOKO1|0140"]["Aba Atribuição"]).select()
        self.apply_parameters_dict("CO07", parameters)
        self.session.findById("wnd[0]").sendVKey(0)
        
        return self._finalizar_pelo_modo_universal("CO07", mode, evidence_path)
    
    def _find(self, sap_id: str):
        try:
            return self.session.findById(sap_id)
        except:
            pass

        test_id = sap_id
        substituicoes = [
            (":0100", ":0120"), (":0120", ":0100"),
            (":7100", ":7120"), (":7120", ":7100"),
            (":1100", ":1120"), (":1120", ":1100")
        ]
        
        for de, para in substituicoes:
            if de in sap_id:
                try:
                    return self.session.findById(sap_id.replace(de, para))
                except:
                    continue

        try:
            tech_name = sap_id.split('/')[-1].replace("ctxt", "").replace("txt", "").replace("chk", "")
            return self.session.findById("wnd[0]").findByName(tech_name, "")
        except:
            pass

        return self.session.findById(sap_id)

    def apply_parameters_dict(self, tcode: str, params: dict[str, str]) -> tuple[dict[str, str], str, str]:
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
                    obj = self._find(sap_id)
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
            return remaining_params, "", "TAB"

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
    
    def _get_tcode_config(self, tcode: str) -> dict:
        """Busca as configurações 'metadata' da transação no YAML."""
        return self.field_map.get(tcode.upper(), {}).get("metadata", {})

    def _finalizar_pelo_modo_universal(self, tcode: str, mode: str, evidence_path: str) -> SapResult:
        """Centraliza a lógica de salvamento ou simulação para qualquer módulo."""
        config = self._get_tcode_config(tcode)
        is_real = self._is_real_mode(mode)
        
        if is_real:
            if tcode.upper() in ["IW31", "IW34"]:
                try: self.session.findById("wnd[0]/tbar[1]/btn[25]").press(); time.sleep(1.0)
                except: pass
            
            btn_save = config.get("btn_salvar", "wnd[0]/tbar[0]/btn[11]")
            try: self.session.findById(btn_save).press()
            except: self.session.findById("wnd[0]").sendVKey(11)
        else:
            btn_sim = config.get("btn_simular")
            if btn_sim:
                try: self.session.findById(btn_sim).press(); time.sleep(1.0)
                except: pass
            self.session.findById("wnd[0]").sendVKey(0) 

        time.sleep(1.5)
        
        popup_msgs = []
        while self._popup_exists():
            txt = self._popup_text(); 
            if txt: popup_msgs.append(txt)
            self._dismiss_popup(); time.sleep(0.5)

        sb = self._statusbar_text()
        sb_type = self._statusbar_type()
        
        if is_real and sb_type in {"E", "A", "X"}:
            ev = self._capture_error_evidence(evidence_path, "STATUSBAR")
            return SapResult("FAIL", "STATUSBAR", sb or "Erro ao salvar no SAP", ev)

        ev = self._capture_success_evidence(evidence_path)
        msg = sb or f"Processado com sucesso ({mode.upper()})"
        if popup_msgs: msg += f" | POPUP: {' || '.join(popup_msgs)}"
            
        return SapResult("PASS", "OK", msg, ev)

    def run_tcode(
        self,
        tcode: str,
        parameters: dict,
        explanation: str,
        evidence_path: str = "",
        mode: str = "real",
        shared_context: dict = None
    ) -> SapResult:
        try:
            self._ensure_session()
            tcode_u = tcode.upper()
            exec_mode = self._normalize_mode(mode)
            
            parameters = enrich_params(tcode, explanation, "", shared_context=shared_context)
            parameters.update({k: v for k, v in parameters.items() if v}) 
            
            if tcode_u == "IW41":
                return self._run_iw41_flow(parameters, evidence_path, mode=exec_mode)

            if tcode_u in ["IP41", "IP42"]:
                return self._run_ip41_ip42_flow(tcode, parameters, explanation, evidence_path, mode=exec_mode)

            if tcode_u in ["CO01", "CO02"]:
                return self._run_pp_order_flow(tcode, parameters, evidence_path, exec_mode)

            if tcode_u == "CO11N":
                return self._run_pp_co11n_flow(parameters, evidence_path, exec_mode)

            if tcode_u == "CO07":
                return self._run_pp_co07_flow(parameters, evidence_path, exec_mode)
            
            self.open_tcode(tcode)

            params_to_fill = {k: v for k, v in (parameters or {}).items() if v is not None and str(v).strip() != ""}
            popup_msgs = []
            max_telas = 15
            tela_atual = 0

            has_order_operation_flow = tcode_u in {"IW31", "IW34"} and bool(
                self._pending_iw31_operation_fields(params_to_fill)
            )
            order_operations_enter_done = False
            is_iw32_print_flow = tcode_u == "IW32" and "imprimir" in (explanation or "").lower()
            iw21_z4_popup_done = False

            while params_to_fill and tela_atual < max_telas:
                tela_atual += 1
                self._handle_all_popups()
                params_to_fill, error_msg, action_taken = self.apply_parameters_dict(tcode, params_to_fill)
                
                if tcode_u == "IW21" and not iw21_z4_popup_done:
                    popup_handled = self._handle_iw21_z4_popup(parameters)
                    if popup_handled:
                        iw21_z4_popup_done = True

                if has_order_operation_flow:
                    pending_ops = self._pending_iw31_operation_fields(params_to_fill)
                    if not pending_ops and not order_operations_enter_done:
                        self.session.findById("wnd[0]").sendVKey(0)
                        time.sleep(0.8)
                        order_operations_enter_done = True
                        continue

                if error_msg:
                    ev = self._capture_error_evidence(evidence_path, "UNMAPPED_PARAM")
                    return SapResult("FAIL", "UNMAPPED_PARAM", error_msg, ev)

                tentativas_popup = 0
                popups_dismissed = False
                while self._popup_exists() and tentativas_popup < 5:
                    txt = self._popup_text()
                    if txt: popup_msgs.append(txt)
                    self._dismiss_popup()
                    tentativas_popup += 1
                    popups_dismissed = True

                sb_type = self._statusbar_type()
                if sb_type in {"E", "A", "X"}:
                    sb = self._statusbar_text()
                    ev = self._capture_error_evidence(evidence_path, "STATUSBAR")
                    return SapResult("FAIL", "STATUSBAR", sb or "Erro SAP durante preenchimento", ev)

                if params_to_fill:
                    if not popups_dismissed:
                        if action_taken == "TEXT":
                            self.session.findById("wnd[0]").sendVKey(0)
                            time.sleep(0.7)
                        elif action_taken == "NONE":
                            if has_order_operation_flow and order_operations_enter_done:
                                time.sleep(0.3)
                            elif str(self.session.Info.Program).strip() == "SAPLIQS0" and tcode_u == "IW31":
                                self.session.findById("wnd[0]").sendVKey(3)
                                time.sleep(0.7)
                            else:
                                self.session.findById("wnd[0]").sendVKey(0)
                                time.sleep(0.7)

            if params_to_fill:
                tcode_maps = self.field_map.get(tcode_u, {})
                todos_campos_yaml = set()
                for screen_key, fields in tcode_maps.items():
                    if isinstance(fields, dict) and screen_key != "metadata":
                        todos_campos_yaml.update(fields.keys())
                
                campos_yaml_norm = {self._norm_key(c) for c in todos_campos_yaml}
                reais_faltantes = [
                    k for k in params_to_fill.keys() 
                    if self._norm_key(k) in campos_yaml_norm
                ]

                if reais_faltantes:
                    msg = f"Campos obrigatórios mapeados no YAML não encontrados no SAP: {reais_faltantes}"
                    ev = self._capture_error_evidence(evidence_path, "UNMAPPED_PARAM")
                    return SapResult("FAIL", "UNMAPPED_PARAM", msg, ev)
                
            self.execute_default()
            time.sleep(1.5)

            if is_iw32_print_flow and tcode_u == "IW32":
                try:
                    self.session.findById("wnd[0]/tbar[0]/btn[86]").press()
                    time.sleep(1.5)
                    self.session.findById("wnd[1]/usr/tblSAPLIPRTTC_WORKPAPERS").getAbsoluteRow(8).selected = True
                    self.session.findById("wnd[1]/tbar[0]/btn[16]").press()
                    time.sleep(3.0)
                    ev = self._capture_success_evidence(evidence_path)
                    msg = "Visualização de impressão gerada com sucesso."
                    try:
                        self.session.findById("wnd[0]/tbar[0]/btn[12]").press()
                        time.sleep(0.8)
                        if self._popup_exists(): self._dismiss_popup()
                        self.session.findById("wnd[0]/tbar[0]/btn[12]").press()
                    except: pass
                    return SapResult("PASS", "OK", f"{msg} ({exec_mode.upper()})", ev)
                except Exception as e:
                    ev = self._capture_error_evidence(evidence_path, "STATUSBAR")
                    return SapResult("FAIL", "EXCEPTION", f"Falha na impressão: {e}", ev)

            while self._popup_exists():
                txt = self._popup_text()
                if txt: popup_msgs.append(txt)
                self._dismiss_popup()
                time.sleep(0.5)

            tentativas_limpeza = 0
            while self._statusbar_type() in {"W", "I", "S"} and tentativas_limpeza < 3:
                try:
                    self.session.findById("wnd[0]").sendVKey(0)
                    time.sleep(0.7)
                except: pass
                tentativas_limpeza += 1

            return self._finalizar_pelo_modo_universal(tcode, exec_mode, evidence_path)

        except Exception as e:
            dump = dump_screen(self.session) if self.session else ""
            ev = self._capture_error_evidence(evidence_path, "STATUSBAR")
            return SapResult("FAIL", "EXCEPTION", f"{e} | DUMP: {dump}", ev)
        
        finally: 
            self.go_to_initial_screen()
    
    def _handle_all_popups(self):
        """Trata popups de forma agressiva (Data, Avisos, Confirmações)"""
        if not self.session:
            return False

        try:
            wnd1 = self.session.findById("wnd[1]", False)
            if wnd1:                
                botoes = [
                    "tbar[0]/btn[0]",    
                    "usr/btnBUTTON_1",   
                    "tbar[0]/btn[11]",   
                    "usr/btnSPOP-OPTION1" 
                ]
                
                for btn_path in botoes:
                    try:
                        wnd1.findById(btn_path).press()
                        time.sleep(0.8)
                        return True
                    except:
                        continue
                
                wnd1.sendVKey(0)
                time.sleep(0.8)
                return True
        except Exception:
            pass
        return False