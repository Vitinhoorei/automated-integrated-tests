# src/params_parser.py
from __future__ import annotations


def parse_parameters(raw: str) -> dict[str, str]:
    """
    Formato aceito (o teu):
      "LI: D10167100-S10167120 | Grupo Planejamento: MUT | Campo Ordenação: Manut Utilitários"

    - separa por "|"
    - cada item vira chave:valor (split no primeiro ":")
    """
    raw = (raw or "").strip()
    if not raw:
        return {}

    parts = [p.strip() for p in raw.split("|") if p.strip()]
    out: dict[str, str] = {}

    for p in parts:
        if ":" not in p:
            continue
        k, v = p.split(":", 1)
        k = k.strip()
        v = v.strip()
        if k:
            out[k] = v

    return out
