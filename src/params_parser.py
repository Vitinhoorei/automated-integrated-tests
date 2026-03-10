from __future__ import annotations
import unicodedata

PARAM_ALIASES = {
    "LI": "Local de instalação",
    "LOCAL INSTALAÇÃO": "Local de instalação",
    "LOCAL DE INSTALAÇÃO": "Local de instalação",

    "Nº EQUIPAMENTO": "Nº equipamento",
    "EQUIPAMENTO": "Nº equipamento",
    "EQUIPAM": "Nº equipamento",
    "N EQUIPAMENTO": "Nº equipamento",

    "PRIORIDADE": "Prioridade",

    "TRABALHO": "Trabalho",
    "TRAB": "Trabalho",

    "Nº COLABORADORES": "Nº colaboradores",
    "NUMERO COLABORADORES": "Nº colaboradores",
    "NUM COLABORADORES": "Nº colaboradores",
    "N COLABORADORES": "Nº colaboradores",
    "N COLAB": "Nº colaboradores",
    
    "IMPRIMIR": "Imprimir",

    "TIPOATVMNT": "Tipo de atividade de manutenção",
    "TIPO ATVMNT": "Tipo de atividade de manutenção",
    "TIPO ATIVIDADE MANUTENÇÃO": "Tipo de atividade de manutenção",
    
    "RAMAL": "Ramal",
    "CRIAR NOTA": "Criar Nota",
    "PRIORIDADE DA NOTA": "Prioridade da Nota",
    "NOTIFICADOR": "Notificador",
}

def remove_acento(txt: str) -> str:
    return ''.join(
        c for c in unicodedata.normalize('NFD', txt)
        if unicodedata.category(c) != 'Mn'
    )

def normalize_key(key: str) -> str:
    """
    Normaliza chave do Excel para o nome esperado no field_map.yaml
    """
    k = key.strip().upper()
    if k in PARAM_ALIASES:
        return PARAM_ALIASES[k]
    
    k_sem_acento = remove_acento(k)

    if k_sem_acento in PARAM_ALIASES:
        return PARAM_ALIASES[k_sem_acento]
    return key.strip()

def parse_parameters(raw: str) -> dict[str, str]:
    """
    Formato aceito:
      "LI: D10167100-S10167120 | Grupo Planejamento: MUT | Campo Ordenação: Manut Utilitários"

    - separa por "|"
    - cada item vira chave:valor
    - aplica alias para nomes conhecidos
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
