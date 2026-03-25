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

        if tcode_u in ["IW32", "IW41", "CO02", "CO11N"]:
            if "Ordem" not in params and "Ordem" in self.shared_context:
                params["Ordem"] = self.shared_context["Ordem"]
                print(f"[CTX-DEBUG] Ordem reutilizada do contexto para {tcode_u}: {params['Ordem']}")

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

    def analisar_erro_sap(self, tcode, status_message, dump_path=None, params=None):
        normalized = self._normalize_error(tcode, status_message, params)
        saved = self.repo.get(normalized)
        if saved:
            return {
                "causa_raiz": saved["causa_raiz"],
                "sugestao_correcao": saved["sugestao_correcao"],
                "parametro_sugerido": saved.get("parametro_sugerido"),
                "confianca": saved.get("confianca", 80),
                "justificativa": "Erro conhecido na base."
            }

        contexto_conhecimento = self.rules_provider.obter_contexto_relevante(params or {})
        dump_text = self._read_dump(dump_path)
        campos_disponiveis = []

        try:
            with open("configs/field_map.yaml", "r", encoding="utf-8") as f:
                field_map = yaml.safe_load(f)
            if tcode in field_map:
                for screen in field_map[tcode]:
                    campos_disponiveis.extend(field_map[tcode][screen].keys())
        except Exception:
            pass

        prompt = f"""
                    Você é um especialista em SAP (PM/PP). O teste falhou.
                    TCODE: {tcode}
                    MENSAGEM SAP: {status_message}

                    BASE DE CONHECIMENTO (VALORES VÁLIDOS):
                    {contexto_conhecimento}

                    VALORES ENVIADOS NA PLANILHA:
                    {params}

                    CAMPOS QUE O ROBÔ RECONHECE NA TELA:
                    {campos_disponiveis}

                    INSTRUÇÕES:
                    1. Identifique qual campo está vazio ("") nos 'Valores Enviados'.
                    2. Se for 'TipoAtvMnt', escolha um CÓDIGO NUMÉRICO na 'Base de Conhecimento' para o tipo de ordem do teste com base naquele teste que esta sendo executado.
                    3. Se for 'Prioridade' e já tiver valor (ex: 2), NÃO ALTERE.
                    4. IMPORTANTE: O nome do campo no retorno deve ser EXATAMENTE igual a um dos nomes da lista 'CAMPOS QUE O ROBÔ RECONHECE NA TELA'.
                    5. Se na planilha está 'TipoAtvMnt' e na tela está 'Tipo de atividade de manutenção', use o nome da tela.
                    6. Responda apenas o código técnico. Ex: Tipo de atividade de manutenção=26.

                    Retorne APENAS JSON:
                    {{
                    "causa_raiz": "O campo X estava vazio",
                    "sugestao_correcao": "Preencher com código Y",
                    "parametro_sugerido": "NOME_DO_CAMPO_DA_TELA=CODIGO_ESCOLHIDO",
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
        self.repo.save(normalized, result)

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
