from pathlib import Path
from datetime import datetime
import re

def ensure_dir(path: str) -> None:
    """
    Garante que o diretório exista.
    """
    Path(path).mkdir(parents=True, exist_ok=True)

def _sanitize(text: str) -> str:
    """
    Normaliza texto para uso seguro em nomes de arquivos.
    """
    if not text:
        return "UNK"
    text = str(text)
    text = re.sub(r"[^\w\-]+", "_", text.strip())
    return text[:60] or "UNK"

def evidence_filename(
    exec_id: str,
    sheet: str,
    row: int,
    tcode: str,
    status: str,
    source: str | None = None,
) -> str:
    """
    Gera nome padronizado e informativo para evidência.

    Exemplo:
    EXEC123_PM_r45_IW31_FAIL_STATUSBAR_20260304_142233.png
    """
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    safe_exec = _sanitize(exec_id)
    safe_sheet = _sanitize(sheet)
    safe_row = f"r{row}"
    safe_tcode = _sanitize(tcode or "NO_TCODE")
    safe_status = _sanitize(status or "UNK")
    safe_source = _sanitize(source) if source else None

    parts = [
        safe_exec,
        safe_sheet,
        safe_row,
        safe_tcode,
        safe_status,
    ]

    if safe_source:
        parts.append(safe_source)

    parts.append(ts)

    return "_".join(parts) + ".png"