import re
import unicodedata
from params_parser import parse_parameters

TIPO_RE = re.compile(r"\btipo\s+([A-Z0-9_]+)\b", re.IGNORECASE)

def _norm_key(text: str) -> str:
    text = (text or "").strip().lower()
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return " ".join(text.split())

def enrich_params(tcode: str, explanation: str, raw_param: str, shared_context: dict = None) -> dict[str, str]:
    params = parse_parameters(raw_param)
    m = TIPO_RE.search(explanation or "")
    tipo = m.group(1).upper() if m else None
    tcode_u = (tcode or "").upper().strip()
    expl_l = (explanation or "").lower()

    if tcode_u in ["CO02", "CO11N"] and not params.get("Ordem"):
        if shared_context and shared_context.get("Ordem"):
            params["Ordem"] = shared_context.get("Ordem")

    alias_to_target = {
        # PM
        "li": "LI", 
        "local de instalacao": "Local de instalação", 
        "nº equipamento": "Nº equipamento",
        "n equipamento": "Nº equipamento", 
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
        "unidade": "Unidade do ciclo",
        "unidade duracao": "Unidade duração", 
        "unidade trabalho": "Unidade trabalho", 
        "grupo planejamento": "Grupo de planejamento",
        "status plano": "Status do plano", 
        "utilizacao": "Utilização", 
        "ctg plano de manutencao": "Ctg.plano de manutenção",
        "descricao da operacao": "Descrição da operação", 
        "campo ordenacao": "Campo Ordenação",
        "estrategia de manutencao": "Estratégia", 
        "estrategia": "Estratégia",

        # PP
        "material": "Material", 
        "centro de producao": "Centro de Produção", 
        "centro": "Centro de Produção",            
        "qtd total": "Quantidade", 
        "quantidade total": "Quantidade", 
        "quantidade": "Quantidade",
        "tipo de programação": "Tipo de programação", 
        "texto breve": "Descrição", 
        "unidade de medida": "Unidade de medida",
        "data base inicio": "Data início", 
        "data inicio": "Data início", 
        "data fim": "Data fim",
        "centro de lucro": "Centro de lucro", 
        "centro lucro": "Centro de lucro", 
        "operacao": "Operação", 
        "numero da ordem": "Ordem"
    }

    normalized: dict[str, str] = {}
    for key, value in params.items():
        key_norm = _norm_key(key)
        target = alias_to_target.get(key_norm, key.strip())
        normalized[target] = str(value).strip()

    params = normalized

    if tcode_u == "IW21" and tipo:
        params["Tipo de nota"] = tipo
    elif tcode_u in ["IW31", "IW34"] and tipo:
        params["Tipo de ordem"] = tipo

    if "LI" in params:
        li_value = params.pop("LI")
        params["Local de instalação"] = li_value

    if tcode_u in ["IW31", "IW34"]:
        has_op = any("Trabalho" in k or "colaborador" in k.lower() or "Operação" in k or "operação" in k.lower() for k in params)
        if has_op: params["Aba Operações"] = "X"
            
    if tcode_u == "IW31":
        if "Nota" not in params or not str(params.get("Nota", "")).strip():
            params["Criar Nota"] = "X"
            if "Prioridade" in params: params["Prioridade da Nota"] = params["Prioridade"]
            params.setdefault("Notificador", "teste"); params.setdefault("Ramal", "1111")
        for k in list(params.keys()):
            if k.strip().lower() == "nota": params.pop(k)
    
    if tcode_u == "IP41":
        val_prog = str(params.get("Campo Ordenação", "")).strip().lower()
        if val_prog in ["manut utilitários", "manut utilitarios", "manut. utilitários"]:
            params["Campo Ordenação"] = "M8"

    if tcode_u in ["CO01", "CO02", "CO07", "CO11N"] and tipo:
        params["Tipo de ordem"] = tipo

    if tcode_u == "CO02":
        if "data" in expl_l or "modificar" in expl_l: params["_FLUXO"] = "DATAS"
        elif "liberar" in expl_l or "imprimir" in expl_l: params["_FLUXO"] = "LIBERAR_IMPRIMIR"
        elif "encerramento" in expl_l or "teco" in expl_l or "concluir" in expl_l: params["_FLUXO"] = "TECO"

    if tcode_u == "CO11N":
        op_val = str(params.get("Operação", "")).strip()
        if "," in op_val: params["_LISTA_OPERACOES"] = [o.strip() for o in op_val.split(",")]
        else: params["_LISTA_OPERACOES"] = [op_val] if op_val else ["0010"] 
    
    if tcode_u == "CO07":
        val_prog = str(params.get("Tipo de programação", "")).lower()
        if "dia" in val_prog or "frente" in val_prog:
            params["Tipo de programação"] = "4"

    return params