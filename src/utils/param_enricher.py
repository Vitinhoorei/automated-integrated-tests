import re
from .params_parser import parse_parameters

TIPO_RE = re.compile(r"\btipo\s+([A-Z0-9_]+)\b", re.IGNORECASE)

def enrich_params(tcode: str, explanation: str, raw_params: str) -> dict:
    params = parse_parameters(raw_params)
    tcode_u = tcode.upper().strip()
    
    m = TIPO_RE.search(explanation or "")
    tipo_detectado = m.group(1).upper() if m else None

    if tcode_u == "IW21" and tipo_detectado:
        params.setdefault("Tipo de nota", tipo_detectado)

    elif tcode_u in ["IW31", "IW34"]:
        if tipo_detectado:
            params.setdefault("Tipo de ordem", tipo_detectado)
        
        has_op = any("trabalho" in k.lower() or "colaborador" in k.lower() for k in params)
        if has_op:
            params["Aba Operações"] = "X"

        if "Nota" not in params:
            params["Criar Nota"] = "X"
            params.setdefault("Notificador", "AUTO_BOT")
            params.setdefault("Ramal", "1111")

    return params