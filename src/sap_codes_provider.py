import yaml
import os

class SAPRulesProvider:
    def __init__(self, path="configs/sap_rules.yaml"):
        self.path = path
        self.rules = self._load()

    def _load(self):
        if os.path.exists(self.path):
            with open(self.path, "r", encoding="utf-8") as f:
                return yaml.safe_load(f)
        return {}

    def obter_contexto_relevante(self, params):
        if not self.rules or not params:
            return ""

        contexto_str = "\n--- BASE DE CONHECIMENTO SAP (REGRAS DE NEGÓCIO) ---\n"
        tipo_ordem_atual = None
        for k, v in params.items():
            if "TIPO" in k.upper() and "ORDEM" in k.upper():
                tipo_ordem_atual = str(v).upper().strip()
                break
        
        regras_raiz = self.rules.get("REGRAS_DE_CONTEXTO", {})
        mapa_atividades = regras_raiz.get("TIPO_ORDEM_X_ATIVIDADE", {})

        if tipo_ordem_atual in mapa_atividades:
            dados = mapa_atividades[tipo_ordem_atual]
            contexto_str += f"CAMPO 'TipoAtvMnt' para Ordem {tipo_ordem_atual}:\n"
            for item in dados.get("valores_validos", []):
                contexto_str += f"  - Código: {item['codigo']} ({item['nome']})\n"
        
        prioridades = regras_raiz.get("PRIORIDADES", [])
        if prioridades:
            contexto_str += f"CAMPO 'Prioridade' (Valores aceitos):\n"
            for p in prioridades:
                contexto_str += f"  - Código: {p['codigo']} ({p['nome']})\n"

        contexto_str += "\nIMPORTANTE: Use APENAS o CÓDIGO numérico no campo sugerido.\n"
        contexto_str += "-----------------------------------------------------\n"
        
        return contexto_str