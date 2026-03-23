from __future__ import annotations
import unicodedata
import re
import yaml
from pathlib import Path

def load_aliases() -> dict:
    path = Path("configs/aliases.yaml")
    if not path.exists():
        return {}
    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    return {normalize_alias_key(k): v for k, v in data.items()}

def remove_acento(txt: str) -> str:
    return ''.join(c for c in unicodedata.normalize('NFD', txt) if unicodedata.category(c) != 'Mn')

def normalize_alias_key(key: str) -> str:
    key = key.strip().upper()
    key = remove_acento(key)
    key = re.sub(r"[^\w\s]", " ", key)
    return " ".join(key.split())

PARAM_ALIASES = load_aliases()

def normalize_key(key: str) -> str:
    alias_key = normalize_alias_key(key)
    return PARAM_ALIASES.get(alias_key, key.strip())

def parse_parameters(raw: str) -> dict[str, str]:
    raw = (raw or "").strip()
    if not raw: return {}
    parts = [p.strip() for p in raw.split("|") if p.strip()]
    out: dict[str, str] = {}
    for p in parts:
        if ":" not in p: continue
        k, v = p.split(":", 1)
        out[normalize_key(k)] = v.strip()
    return out