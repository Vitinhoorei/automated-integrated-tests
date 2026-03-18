from __future__ import annotations
import re
import shutil
import uuid
from pathlib import Path

from src.config import AppConfig
from src.utils.evidence import ensure_dir, evidence_filename
from src.services.ai_integrator import AITestIntegrator
from src.services.excel_manager import ExcelManager 
from src.core.sap_driver import SapDriver
from src.handlers.iw31_handler import IW31Handler
from src.handlers.iw41_handler import IW41Handler
from src.handlers.iw21_handler import IW21Handler
from handlers.iw32_handler import IW32Handler
from src.handlers.ip_handler import IPHandler

def _copy_to_output(original_path: str, output_dir: str) -> str:
    ensure_dir(output_dir)
    src = Path(original_path)
    dst = Path(output_dir) / src.name
    shutil.copy2(src, dst)
    return str(dst)

def run_excel_tests(xlsx_path: str, sheet_name: str, make_copy: bool = True) -> str:
    cfg = AppConfig()
    ensure_dir(cfg.evidence_dir)
    exec_id = uuid.uuid4().hex[:10]

    driver = SapDriver()
    driver.connect()
    ai = AITestIntegrator()

    handlers = {
        "IW31": IW31Handler(driver),
        "IW34": IW31Handler(driver),
        "IW41": IW41Handler(driver),
        "IW21": IW21Handler(driver),
        "IW32": IW32Handler(driver),
        "IP41": IPHandler(driver),
        "IP42": IPHandler(driver),
    }

    work_xlsx = _copy_to_output(xlsx_path, cfg.output_dir) if make_copy else xlsx_path

    with ExcelManager(work_xlsx) as excel:
        if not sheet_name or sheet_name.upper() == "ALL":
            target_sheets = excel.get_sheets()
        else:
            target_sheets = [s.strip() for s in sheet_name.split(",") if s.strip()]

        for sname in target_sheets:
            try:
                rows = excel.read_rows(sname)
            except Exception as e:
                print(f"[SKIP] Erro ao ler aba {sname}: {e}")
                continue

            for item in rows:
                if str(item.mode).lower() not in ["executar", "simulado"]:
                    continue

                ev_name = evidence_filename(exec_id, item.sheet, item.index, item.tcode, "RUN")
                ev_path = str(Path(cfg.evidence_dir) / ev_name)

                handler = handlers.get(item.tcode.upper())
                if not handler:
                    print(f"⚠️ Sem handler para {item.tcode}")
                    continue

                from src.utils.param_enricher import enrich_params
                params = enrich_params(item.tcode, item.explanation, item.params)
                result = handler.run(item.tcode, params, ev_path, mode=item.mode)
                excel.write_result(item.sheet, item.index, result)
                print(f"[{item.sheet} r{item.index}] {item.tcode} -> {result.status}")

    return work_xlsx