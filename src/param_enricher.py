import re
import unicodedata

from params_parser import parse_parameters

TIPO_RE = re.compile(r"\btipo\s+([A-Z0-9_]+)\b", re.IGNORECASE)


def _norm_key(text: str) -> str:
    text = (text or "").strip().lower()
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return " ".join(text.split())


def enrich_params(tcode: str, explanation: str, raw_param: str) -> dict[str, str]:
    params = parse_parameters(raw_param)

    # extrai tipo do texto "tipo Z1", "tipo ZCOR" etc.
    m = TIPO_RE.search(explanation or "")
    tipo = m.group(1).upper() if m else None

    tcode_u = (tcode or "").upper().strip()

    # aplica regra por transacao
    if tcode_u == "IW21" and tipo:
        params["Tipo de nota"] = tipo
    elif tcode_u == "IW31" and tipo:
        params["Tipo de ordem"] = tipo

    # Normaliza aliases para reduzir divergencias de escrita/acentuacao.
    alias_to_target = {
        "li": "LI",
        "local de instalacao": "Local de instalação",
        "plano manutencao": "Plano manutenção",
        "ctg.plano manut.": "Ctg.plano manut.",
        "ctg plano manut.": "Ctg.plano manut.",
        "grupo planejamento": "Grupo Planejamento",
        "campo ordenacao": "Campo Ordenação",
        "estrategia": "Estratégia",
        "tipo de nota": "Tipo de nota",
        "tipo de ordem": "Tipo de ordem",
    }

    normalized: dict[str, str] = {}
    for key, value in params.items():
        key_norm = _norm_key(key)
        target = alias_to_target.get(key_norm, key.strip())
        normalized[target] = str(value).strip()
    params = normalized

    # Regra geral: LI -> Local de instalacao (exceto IW21 na tela inicial).
    if "LI" in params:
        li_value = params.pop("LI")
        if tcode_u == "IW21":
            pass
        elif tcode_u in {"IP41", "IP42"}:
            params["Plano manutenção"] = li_value
        else:
            params["Local de instalação"] = li_value

    # Em IP41/IP42, se vier Local de instalação explicitamente, mapeia para Plano manutenção.
    if tcode_u in {"IP41", "IP42"} and "Local de instalação" in params:
        params["Plano manutenção"] = params.pop("Local de instalação")

    # Valores fixos solicitados para os dois casos problemáticos.
    if tcode_u in {"IP41", "IP42"}:
        params["Ctg.plano manut."] = "Estudo de estabilidade"
        params.pop("Grupo Planejamento", None)

    if tcode_u == "IP42":
        params["Estratégia"] = "EF001"
        params.pop("Campo Ordenação", None)
    else:
        params.pop("Campo Ordenação", None)
        params.pop("Estratégia", None)

    return params
