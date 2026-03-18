import os
import re
from pathlib import Path
from datetime import datetime

def ensure_dir(path: str) -> None:
    """Garante que o diretório exista."""
    Path(path).mkdir(parents=True, exist_ok=True)

def _sanitize(text: str) -> str:
    """Normaliza texto para uso seguro em nomes de arquivos."""
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
) -> str:
    """Gera nome padronizado para evidência PNG."""
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    parts = [
        _sanitize(exec_id),
        _sanitize(sheet),
        f"r{row}",
        _sanitize(tcode or "NO_TCODE"),
        _sanitize(status or "UNK"),
        ts
    ]

    return "_".join(parts) + ".png"