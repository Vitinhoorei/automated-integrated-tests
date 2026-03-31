import os
import re
import json
import hashlib
import requests
import config
import yaml

from param_enricher import enrich_params
from error_repository import ErrorRepository
from sap_codes_provider import SAPRulesProvider 

class AITestIntegrator:

    def __init__(self):
        self.api_url = getattr(config, "IA_BASE_URL", os.getenv("IA_BASE_URL", ""))
        self.api_key = getattr(config, "IA_API_KEY", os.getenv("IA_API_KEY", ""))
        self.shared_context = {}
        self.historico_sucesso = []
        self.repo = ErrorRepository()
        self.rules_provider = SAPRulesProvider()

    def _chamar_ia(self, prompt: str, json_mode: bool = False):
        payload = {
            "question": prompt,
            "model": "gpt-4o-mini",
            "temperature": 0
        }

        headers = {
            "Content-Type": "application/json",
            "apiKey": self.api_key
        }

        try:
            resp = requests.post(
                self.api_url,
                headers=headers,
                json=payload,
                timeout=90
            )

            if resp.status_code != 200:
                return "{}" if json_mode else "Erro IA"

            text = resp.json().get("response", "")

            if json_mode:
                match = re.search(r"\{.*\}", text, re.DOTALL)
                if match:
                    return match.group(0)
                return "{}"
            return text

        except Exception:
            return "{}" if json_mode else "Falha IA"

    def preparar_parametros(self, tcode, explanation, raw_params):
        params = enrich_params(tcode, explanation, raw_params)
        tcode_u = tcode.upper().strip()
        
        if tcode_u == "IW31":
            if "Nota" not in params and "Nota" in self.shared_context:
                params["Nota"] = self.shared_context["Nota"]

        if tcode_u in ["IW32", "IW41"]:
            if "Ordem" not in params and "Ordem" in self.shared_context:
                params["Ordem"] = self.shared_context["Ordem"]

        return params

    def extrair_id_integrado(self, tcode, status_message):
        if not status_message:
            return

        prompt = f"""
                    Extraia entidade e ID da mensagem SAP:
                    "{status_message}"
                    
                    Retorne JSON:
                    {{"entidade":"NOME","id":"VALOR"}}
                """

        resposta = self._chamar_ia(prompt, json_mode=True)

        try:
            data = json.loads(resposta)
            if data.get("entidade") and data.get("id"):
                self.shared_context[data["entidade"]] = str(data["id"])
        except Exception:
            pass
    
    def _buscar_regra_local(self, tcode, status_message, params):
        if not status_message: return None
            
        msg_lower = status_message.lower()
        tcode_upper = tcode.upper().strip()
        erro_generico = "obrigatório" in msg_lower or "mandatory" in msg_lower
        
        for erro_mapeado, regra in self.repo.errors.items():
            if len(erro_mapeado) == 32: continue 

            tcodes_permitidos = regra.get("tcodes", [])
            if tcodes_permitidos and tcode_upper not in tcodes_permitidos: continue

            campo = regra.get("campo_sugerido", "")
            mensagem_especifica = f"'{campo.lower()}'" in msg_lower
            
            if campo.lower() == "prioridade" and "prioridade da nota" in msg_lower:
                continue 

            if erro_mapeado.lower() in msg_lower or mensagem_especifica:
                valor = regra.get("valor_padrao", "")
                regra_dinamica = regra.get("usar_regra_dinamica")
                
                if regra_dinamica == "TIPO_ORDEM_X_ATIVIDADE":
                    tipo_ordem = params.get("Tipo de ordem", params.get("Tipo de nota", "")).upper()
                    try:
                        with open("sap_codes.yaml", "r", encoding="utf-8") as f:
                            sap_codes = yaml.safe_load(f)
                        regras_atividade = sap_codes.get("REGRAS_DE_CONTEXTO", {}).get("TIPO_ORDEM_X_ATIVIDADE", {})
                        if tipo_ordem in regras_atividade:
                            valor = regras_atividade[tipo_ordem]["valores_validos"][0]["codigo"]
                    except Exception: pass 
                
                if campo and valor:
                    return {
                        "causa_raiz": f"Erro específico identificado: '{campo}'",
                        "sugestao_correcao": f"Preencher '{campo}' com '{valor}'",
                        "parametro_sugerido": f"{campo}={valor}",
                        "confianca": 100,
                        "justificativa": regra.get("justificativa", "Auto-Cura local via Regras YAML/JSON.")
                    }

        if erro_generico:
            for erro_mapeado, regra in self.repo.errors.items():
                if len(erro_mapeado) == 32: continue 

                tcodes_permitidos = regra.get("tcodes", [])
                if tcodes_permitidos and tcode_upper not in tcodes_permitidos: continue

                campo = regra.get("campo_sugerido", "")
                campo_vazio_na_planilha = campo not in params or str(params.get(campo, "")).strip() == ""
                
                if campo_vazio_na_planilha:
                    valor = regra.get("valor_padrao", "")
                    regra_dinamica = regra.get("usar_regra_dinamica")
                    
                    if regra_dinamica == "TIPO_ORDEM_X_ATIVIDADE":
                        tipo_ordem = params.get("Tipo de ordem", params.get("Tipo de nota", "")).upper()
                        try:
                            with open("sap_codes.yaml", "r", encoding="utf-8") as f:
                                sap_codes = yaml.safe_load(f)
                            regras_atividade = sap_codes.get("REGRAS_DE_CONTEXTO", {}).get("TIPO_ORDEM_X_ATIVIDADE", {})
                            if tipo_ordem in regras_atividade:
                                valor = regras_atividade[tipo_ordem]["valores_validos"][0]["codigo"]
                        except Exception: pass 
                    
                    if campo and valor:
                        return {
                            "causa_raiz": "Erro genérico mapeado para campo faltante",
                            "sugestao_correcao": f"Preencher '{campo}' com '{valor}'",
                            "parametro_sugerido": f"{campo}={valor}",
                            "confianca": 100,
                            "justificativa": regra.get("justificativa", "Auto-Cura local genérica.")
                        }

        return None

    def analisar_erro_sap(self, tcode, status_message, dump_path=None, params=None):
        regra_local = self._buscar_regra_local(tcode, status_message, params or {})

        if status_message and ("HHMANU" in status_message or "preço" in status_message.lower()):
            return {
                "causa_raiz": status_message,  # <--- AQUI VAI O TEXTO CRU E EXATO!
                "sugestao_correcao": "Verificar a carga de tarifas na transação KP26 neste ambiente.",
                "parametro_sugerido": "",
                "confianca": 100,
                "justificativa": "Mensagem exata capturada diretamente da tela de Log (ALV Grid)."
            }
        
        if regra_local:
            return regra_local

        normalized = self._normalize_error(tcode, status_message, params)
        saved = self.repo.get(normalized)
        if saved and isinstance(saved, dict) and "causa_raiz" in saved:
            return {
                "causa_raiz": saved["causa_raiz"],
                "sugestao_correcao": saved["sugestao_correcao"],
                "parametro_sugerido": saved.get("parametro_sugerido"),
                "confianca": saved.get("confianca", 80),
                "justificativa": "Erro conhecido no cache da IA."
            }

        contexto_conhecimento = self.rules_provider.obter_contexto_relevante(params or {})
        dump_text = self._read_dump(dump_path)
        campos_disponiveis = []

        try:
            with open("configs/field_map.yaml", "r", encoding="utf-8") as f:
                field_map = yaml.safe_load(f)
            if tcode in field_map:
                for screen in field_map[tcode]:
                    if isinstance(field_map[tcode][screen], dict):
                        campos_disponiveis.extend(field_map[tcode][screen].keys())
        except Exception:
            pass

        prompt = f"""
                    Você é um especialista em SAP PM. O teste falhou.
                    TCODE: {tcode}
                    MENSAGEM SAP: {status_message}

                    BASE DE CONHECIMENTO (VALORES VÁLIDOS):
                    {contexto_conhecimento}

                    VALORES ENVIADOS NA PLANILHA:
                    {params}

                    CAMPOS QUE O ROBÔ RECONHECE NA TELA:
                    {campos_disponiveis}

                    INSTRUÇÕES (SIGA ESTRITAMENTE):
                    1. Identifique o erro com base APENAS na MENSAGEM SAP.
                    2. REGRA DE OURO (PROIBIDO ALUCINAR): Se a MENSAGEM SAP indicar que um dado digitado na planilha não existe, é inválido ou não está cadastrado (Ex: "não existente em EQUI", "Centro não cadastrado", "Material inválido"), VOCÊ NÃO PODE ADIVINHAR O VALOR.
                    3. Nesses casos de dados mestres errados, retorne "parametro_sugerido": "" (string vazia). Apenas explique ao usuário que ele digitou errado na "sugestao_correcao".
                    4. Só preencha o "parametro_sugerido" (formato NOME=VALOR) se for um erro de campo vazio e você tiver certeza da correção baseada na Base de Conhecimento.
                    5. Retorne APENAS um JSON válido.

                    Retorne APENAS JSON neste formato:
                    {{
                    "causa_raiz": "O campo X estava incorreto/vazio.",
                    "sugestao_correcao": "Verificar se o equipamento digitado existe no SAP.",
                    "parametro_sugerido": "",
                    "confianca": 100
                    }}
                """

        resposta = self._chamar_ia(prompt, json_mode=True)
        try:
            data = json.loads(resposta)
        except Exception:
            data = {}

        result = {
            "causa_raiz": data.get("causa_raiz", status_message),
            "sugestao_correcao": data.get("sugestao_correcao", ""),
            "parametro_sugerido": data.get("parametro_sugerido"),
            "confianca": data.get("confianca", 50),
            "justificativa": "Análise IA orientada por Base de Conhecimento e Mapeamento de Tela"
        }
                
        return result
    
    def aplicar_correcao_parametros(self, params, parametro_sugerido):
        if not parametro_sugerido:
            return params
        try:
            if "CODIGO" in parametro_sugerido.upper() or "VALOR" in parametro_sugerido.upper():
                return params

            chave, valor = parametro_sugerido.split("=")
            params[chave.strip()] = valor.strip()
        except Exception:
            pass
        return params

    def _normalize_error(self, tcode, message, params=None):
        msg_limpa = message.lower().strip()
        msg_limpa = re.sub(r'\d+', 'X', msg_limpa)
        msg_limpa = re.sub(r"'.*?'", "'X'", msg_limpa)
        
        contexto = ""
        if params:
            for k, v in params.items():
                if "TIPO" in k.upper() and "ORDEM" in k.upper():
                    contexto = f"|TIPO:{v}"
                    break

        base = f"{tcode}|{msg_limpa}{contexto}"
        return hashlib.md5(base.encode()).hexdigest()

    def _read_dump(self, dump_path):
        if dump_path and os.path.exists(dump_path):
            try:
                with open(dump_path, "r", encoding="utf-8") as f:
                    return f.read()[:2000]
            except Exception:
                pass
        return "Sem dump."