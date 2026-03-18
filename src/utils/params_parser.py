from __future__ import annotations
from .text_normalizer import normalize_text

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

    "ESTRATÉGIA DE MANUTENÇÃO": "Estratégia",
    "ESTRATEGIA DE MANUTENCAO": "Estratégia",
    "ESTRATEGIA": "Estratégia",
}

def parse_parameters(raw_string: str) -> dict:
    if not raw_string or str(raw_string).strip() == "":
        return {}
    
    parts = [p.strip() for p in str(raw_string).split("|") if p.strip()]
    out = {}

    for p in parts:
        if ":" not in p: continue
        k, v = p.split(":", 1)
        
        key_norm = normalize_text(k)
        final_key = PARAM_ALIASES.get(key_norm, k.strip())
        
        out[final_key] = v.strip()
    return out