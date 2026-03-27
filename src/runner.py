from __future__ import annotations

import re
import shutil
import uuid
from pathlib import Path

from config import AppConfig
from evidence import ensure_dir, evidence_filename
from ia_integrator import AITestIntegrator
from planilha_local import (
    format_output_sheet,
    list_sheet_names,
    read_rows,
    write_status_with_fix_details,
)
from sap_automation import SapAutomation
from sap_screen_dump import dump_screen


def _copy_to_output(original_path: str, output_dir: str, sheet_name: str) -> str:
    ensure_dir(output_dir)
    src = Path(original_path)
    texto_busca = f"{src.name} {sheet_name}".upper()
    
    if "PP" in texto_busca:
        modulo_nome = "PP"
    elif "PM" in texto_busca:
        modulo_nome = "PM"
    else:
        modulo_nome = "Geral" 

    name = f"Resultados de Automação - {modulo_nome}{src.suffix}"
    dst = Path(output_dir) / name
    shutil.copy2(src, dst)
    return str(dst)

def run_excel_tests(
    xlsx_path: str,
    sheet_name: str,
    make_copy: bool = True
) -> str:

    cfg = AppConfig()
    ensure_dir(cfg.evidence_dir)

    exec_id = uuid.uuid4().hex[:10]
    sap = SapAutomation(field_map_path="configs/field_map.yaml")
    ai = AITestIntegrator()
    work_xlsx = _copy_to_output(xlsx_path, cfg.output_dir, sheet_name) if make_copy else xlsx_path
    raw_sheet = (sheet_name or "").strip()

    if not raw_sheet or raw_sheet.upper() == "ALL":
        target_sheets = list_sheet_names(work_xlsx)
    else:
        target_sheets =[s.strip() for s in raw_sheet.split(",") if s.strip()]

    rows =[]
    processed_sheets = set()

    for sname in target_sheets:
        try:
            rows.extend(read_rows(work_xlsx, sheet_name=sname))
            processed_sheets.add(sname)
        except ValueError as e:
            print(f"[SKIP] Aba '{sname}' ignorada: {e}")

    valid_modes = {"executar", "simulado"}

    for item in rows:
        mode = (item.mode or "").strip().lower()

        if mode not in valid_modes:
            print(
                f"[{item.sheet_name} r{item.row_index}] "
                f"{item.tcode} -> SKIP | modo vazio ou inválido: '{item.mode}'"
            )
            continue

        fname = evidence_filename(
            exec_id,
            item.sheet_name,
            item.row_index,
            item.tcode,
            "RUN",
        )

        evidence_path = str(Path(cfg.evidence_dir) / fname)

        smart_params = ai.preparar_parametros(
            item.tcode,
            item.explanation,
            item.parameter,
        )

        params = smart_params.copy()
        retry_count = 0
        MAX_RETRY = 3

        while True:

            result = sap.run_tcode(
                item.tcode,
                params,
                item.explanation,
                evidence_path=evidence_path,
                mode=mode,
                shared_context=ai.shared_context
            )
            
            if result.status != "PASS":
                print("DEBUG SAP MESSAGE:", result.message)
            
            # PASS
            if result.status == "PASS":

                ai.extrair_id_integrado(item.tcode, result.message)
                transacao_ok = params.copy()
                transacao_ok["_TCODE"] = item.tcode
                ai.historico_sucesso.append(transacao_ok)

                if retry_count > 0:
                    justificativa = f"⚠️ PASSOU COM AUTO-CORREÇÃO (Tentativa {retry_count + 1}). Atualize a planilha com os dados corretos para evitar erros no futuro. Modo: {item.mode}"
                    status_print = f"PASS (Auto-Healed t-{retry_count + 1})"
                else:
                    justificativa = f"Execução concluída sem erros de primeira. Modo: {item.mode}"
                    status_print = "PASS"

                tcode_config = sap._get_tcode_config(item.tcode)
                regex_padrao = tcode_config.get("regex_sucesso", r"(\d+)")
                
                match = re.search(regex_padrao, result.message, re.IGNORECASE)
                if match:
                    id_gerado = match.group(1)
                    ai.shared_context["UltimoID"] = id_gerado
                    
                    if "nota" in result.message.lower() or "aviso" in result.message.lower():
                        ai.shared_context["Nota"] = id_gerado
                    elif "ordem" in result.message.lower():
                        ai.shared_context["Ordem"] = id_gerado
                    
                    print(f"ID Detectado: {id_gerado}")

                write_status_with_fix_details(
                    xlsx_path=work_xlsx,
                    sheet_name=item.sheet_name,
                    row_index=item.row_index,
                    status="PASS",
                    source=result.source,
                    message=result.message,
                    suggested_fix="",
                    fix_confidence=100,
                    fix_justification=justificativa,
                    evidence_path=result.evidence_path,
                )

                print(
                    f"[{item.sheet_name} r{item.row_index}] "
                    f"{item.tcode} ({item.mode}) -> {status_print} | {result.message}"
                )

                break

            # FAIL
            dump_path = ""

            if sap.session:
                try:
                    dump_path = dump_screen(
                        sap.session,
                        out_dir=cfg.evidence_dir
                    )
                except Exception:
                    pass

            analise = ai.analisar_erro_sap(
                item.tcode,
                result.message,
                dump_path,
                params
            )

            causa = analise.get("causa_raiz", result.message)
            sugestao = analise.get("sugestao_correcao", "")
            confianca = analise.get("confianca", 0)
            justificativa = analise.get("justificativa", "")
            parametro = analise.get("parametro_sugerido")

            if (
                retry_count < MAX_RETRY
                and confianca >= 70
                and parametro
                and "=" in parametro
            ):
                chave_sugerida, valor_sugerido = parametro.split("=", 1)
                chave_sugerida = chave_sugerida.strip()
                valor_sugerido = valor_sugerido.strip()

                if params.get(chave_sugerida) == valor_sugerido:
                    print(f"⚠️ A correção '{parametro}' já foi tentada e o erro continuou. Abortando repetição.")
                else:
                    print(f"🔁 Correção automática aplicada: {parametro} | tentativa {retry_count+1}")
                    
                    params = ai.aplicar_correcao_parametros(
                        params,
                        parametro
                    )
                    retry_count += 1
                    continue

            # FAIL DEFINITIVO
            write_status_with_fix_details(
                xlsx_path=work_xlsx,
                sheet_name=item.sheet_name,
                row_index=item.row_index,
                status="FAIL",
                source=result.source,
                message=causa,
                suggested_fix=sugestao,
                fix_confidence=confianca,
                fix_justification=f"{justificativa} | Modo: {mode}",
                evidence_path=result.evidence_path,
            )

            print(
                f"[{item.sheet_name} r{item.row_index}] "
                f"{item.tcode} ({mode}) -> FAIL | {causa} | "
                f"Sugestão: {sugestao} | Confiança: {confianca}"
            )

            if sap.session:
                try:
                    sap.go_to_initial_screen()
                except Exception as e:
                    print(f"[runner] falha ao voltar para tela inicial: {e}")
            break

    for sname in processed_sheets:
        try:
            format_output_sheet(work_xlsx, sname)
        except Exception as e:

            print(
                f"[WARN] Falha ao formatar aba '{sname}': {e}"
            )

    return work_xlsx