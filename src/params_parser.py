from __future__ import annotations
import unicodedata
import re

PARAM_ALIASES = {
    "LI": "Local de instalação",
    "LOCAL INSTALACAO": "Local de instalação",
    "LOCAL DE INSTALACAO": "Local de instalação",

    "N EQUIPAMENTO": "Nº equipamento",
    "NUM EQUIPAMENTO": "Nº equipamento",
    "NUMERO EQUIPAMENTO": "Nº equipamento",
    "EQUIPAMENTO": "Nº equipamento",
    "EQUIPAM": "Nº equipamento",

    "PRIORIDADE": "Prioridade",

    "TRABALHO": "Trabalho",
    "TRAB": "Trabalho",

    "N COLABORADORES": "Nº colaboradores",
    "NUM COLABORADORES": "Nº colaboradores",
    "NUMERO COLABORADORES": "Nº colaboradores",
    "N COLAB": "Nº colaboradores",

    "IMPRIMIR": "Imprimir",

    "TIPOATVMNT": "Tipo de atividade de manutenção",
    "TIPO ATVMNT": "Tipo de atividade de manutenção",
    "TIPO ATIVIDADE MANUTENCAO": "Tipo de atividade de manutenção",

    "TXT BREVE OPERACAO": "Texto breve Operação",
    "TEXTO BREVE OPERACAO": "Texto breve Operação",

    "RAMAL": "Ramal",
    "CRIAR NOTA": "Criar Nota",
    "PRIORIDADE DA NOTA": "Prioridade da Nota",
    "NOTIFICADOR": "Notificador",
    
    "UNIDADE": "Unidade do ciclo",
    "UNIDADE DO CICLO": "Unidade do ciclo",
    "CTG PLANO DE MANUTENCAO": "Ctg.plano de manutenção",
    "TEXTO PLANO DE MANUTENCAO": "Texto do plano de manutenção",
    "DESCRICAO DA OPERACAO": "Descrição da operação",
    "CAMPO ORDENAÇÃO": "Campo seleção p/planos de manutenção",
    
}

def remove_acento(txt: str) -> str:
    return ''.join(
        c for c in unicodedata.normalize('NFD', txt)
        if unicodedata.category(c) != 'Mn'
    )

def normalize_alias_key(key: str) -> str:
    """
    Normaliza chave para busca no alias
    """
    key = key.strip().upper()
    key = remove_acento(key)
    key = re.sub(r"[^\w\s]", " ", key)
    key = " ".join(key.split())
    
    return key

def normalize_key(key: str) -> str:
    """
    Converte chave recebida para o nome esperado no field_map
    """
    
    alias_key = normalize_alias_key(key)

    if alias_key in PARAM_ALIASES:
        return PARAM_ALIASES[alias_key]

    return key.strip()

def parse_parameters(raw: str) -> dict[str, str]:
    """
    Converte string de parâmetros em dict
    Exemplo entrada:
    LI: D10167100 | PRIORIDADE: 2 | TRABALHO: 1
    Saída:

    {
        "Local de instalação": "D10167100",
        "Prioridade": "2",
        "Trabalho": "1"
    }
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
        key = normalize_key(k)
        value = v.strip()

        if key:
            out[key] = value

    return out