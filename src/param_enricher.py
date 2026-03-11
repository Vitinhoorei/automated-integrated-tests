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
    m = TIPO_RE.search(explanation or "")
    tipo = m.group(1).upper() if m else None
    tcode_u = (tcode or "").upper().strip()

    if tcode_u == "IW21" and tipo:
        params["Tipo de nota"] = tipo

    elif tcode_u in ["IW31", "IW34"] and tipo:
        params["Tipo de ordem"] = tipo

    alias_to_target = {
        "li": "LI",
        "local de instalacao": "Local de instalação",
        "tipo de nota": "Tipo de nota",
        "tipo de ordem": "Tipo de ordem",
        "tipo de atividade de manutencao": "Tipo de atividade de manutenção",
        "tipoatvmnt": "Tipo de atividade de manutenção",
        "trabalho": "Trabalho",
        "trabalho 2": "Trabalho 2",
        "trabalho 3": "Trabalho 3",
        "nº colaboradores": "Nº colaboradores",
        "n colaboradores": "Nº colaboradores",
        "no colaboradores": "Nº colaboradores",
        "nº colaboradores 2": "Nº colaboradores 2",
        "nº colaboradores 3": "Nº colaboradores 3",
        "texto operacao": "Texto Operação",
        "texto operacao 2": "Texto Operação 2",
        "texto operacao 3": "Texto Operação 3",
        "descricao operacao": "Descrição da operação",
        "duracao": "Duração",
        "unidade duracao": "Unidade duração",
        "unidade trabalho": "Unidade trabalho",
        "grupo planejamento": "Grupo de planejamento",
        "status plano": "Status do plano",
        "utilizacao": "Utilização",
        "ctg plano manutencao": "Ctg.plano manutenção",

    }

    normalized: dict[str, str] = {}

    for key, value in params.items():
        key_norm = _norm_key(key)
        target = alias_to_target.get(key_norm, key.strip())
        normalized[target] = str(value).strip()

    params = normalized

    if "LI" in params:
        li_value = params.pop("LI")
        params["Local de instalação"] = li_value

    if tcode_u in ["IW31", "IW34"]:

        has_op = any(
            "Trabalho" in k
            or "colaborador" in k.lower()
            or "Operação" in k
            or "operação" in k.lower()
            for k in params
        )

        if has_op:
            params["Aba Operações"] = "X"
            
    if tcode_u == "IW31":

        if "Nota" not in params or not str(params.get("Nota", "")).strip():
            params["Criar Nota"] = "X"
            if "Prioridade" in params:
                params["Prioridade da Nota"] = params["Prioridade"]
            params.setdefault("Notificador", "teste")
            params.setdefault("Ramal", "1111")
            
        for k in list(params.keys()):
            if k.strip().lower() == "nota":
                params.pop(k)

    return params