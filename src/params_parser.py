from __future__ import annotations

PARAM_ALIASES = {
    "LI": "Local de instalação",
    "LOCAL INSTALAÇÃO": "Local de instalação",
    "LOCAL DE INSTALAÇÃO": "Local de instalação",

    "GRUPO PLANEJAMENTO": "Grupo Planejamento",
    "CAMPO ORDENAÇÃO": "Campo Ordenação",

    "TEXTO BREVE": "Texto Breve",
    "NOTIFICADOR": "Notificador",
    "PRIORIDADE": "Prioridade",
    "RAMAL": "Ramal",
}

def normalize_key(key: str) -> str:
    """
    Normaliza chave do Excel para o nome esperado no field_map.yaml
    """
    k = key.strip().upper()

    if k in PARAM_ALIASES:
        return PARAM_ALIASES[k]
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
