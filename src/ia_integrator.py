import os
import re
import json
import hashlib
import requests
import config

from param_enricher import enrich_params
from error_repository import ErrorRepository

class AITestIntegrator:

    def __init__(self):
        self.api_url = getattr(config, "IA_BASE_URL", os.getenv("IA_BASE_URL", ""))
        self.api_key = getattr(config, "IA_API_KEY", os.getenv("IA_API_KEY", ""))
        self.shared_context = {}
        self.historico_sucesso =[]
        self.repo = ErrorRepository()

    # CHAMADA IA
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

    # PREPARAÇÃO DE PARAMETROS
    def preparar_parametros(self, tcode, explanation, raw_params):
        params = enrich_params(tcode, explanation, raw_params)
        tcode_u = tcode.upper().strip()
        
        if tcode_u == "IW31":

            if "Nota" not in params and "Nota" in self.shared_context:
                params["Nota"] = self.shared_context["Nota"]

        # Compartilhar ORDEM
        if tcode_u in ["IW32", "IW41"]:

            if "Ordem" not in params and "Ordem" in self.shared_context:
                params["Ordem"] = self.shared_context["Ordem"]

        return params

    # EXTRAIR IDS GERADOS PELO SAP
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

    # ANALISE DE ERRO SAP
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

        dump_text = self._read_dump(dump_path)
        campos_disponiveis = []

        try:
            import yaml
            with open("configs/field_map.yaml", "r", encoding="utf-8") as f:
                field_map = yaml.safe_load(f)
            if tcode in field_map:
                for screen in field_map[tcode]:
                    campos_disponiveis.extend(field_map[tcode][screen].keys())

        except Exception:
            pass

        # PROMPT IA
        prompt = f"""
                    Analise o erro SAP.

                    TCODE: {tcode}

                    Mensagem SAP:
                    {status_message}

                    Parametros enviados nesta transação (SAGRADOS - NÃO ALTERE):
                    {params}

                    Campos disponíveis na transação atual:
                    {campos_disponiveis}

                    Histórico de Transações Anteriores com Sucesso:
                    {self.historico_sucesso}

                    Dump técnico:
                    {dump_text}

                    REGRAS OBRIGATÓRIAS:
                    1. PARÂMETROS INTOCÁVEIS: NUNCA sugira a alteração de um parâmetro que já veio preenchido (Ex: não mude o Tipo de Ordem).
                    2. COMPARAÇÃO INTELIGENTE (O SEGREDO): Olhe a lista de "Histórico de Transações Anteriores com Sucesso". Se a transação atual for, por exemplo, "Tipo de ordem"="ZMEL" e faltar um campo, procure no Histórico uma transação passada que TAMBÉM era "ZMEL" e copie o valor que funcionou lá. NUNCA misture parâmetros de tipos de ordem diferentes.
                    3. SÓ PREENCHA O VAZIO: Você só pode sugerir valores para campos que estão vazios ou ausentes na transação atual.
                    4. PROIBIDO INVENTAR DADOS MESTRES. Se não achar uma correspondência exata no Histórico que sirva para a transação atual, retorne "confianca": 0 e "parametro_sugerido": null.
                    5. Sugira o valor no formato: Campo=Valor.

                    Retorne APENAS JSON:

                    {{
                    "causa_raiz":"texto explicando o erro",
                    "sugestao_correcao":"texto instruindo o usuário",
                    "parametro_sugerido": "Campo=Valor" ou null,
                    "confianca": 0 a 100
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
            "justificativa": "Análise IA"

        }

        self.repo.save(normalized, result)

        return result

    # APLICAR CORREÇÃO
    def aplicar_correcao_parametros(self, params, parametro_sugerido):

        if not parametro_sugerido:
            return params
        try:
            chave, valor = parametro_sugerido.split("=")
            params[chave.strip()] = valor.strip()

        except Exception:
            pass
        return params

    # NORMALIZA ERRO
    def _normalize_error(self, tcode, message, params=None):
        msg_limpa = message.lower().strip()
        msg_limpa = re.sub(r'\d+', 'X', msg_limpa)
        msg_limpa = re.sub(r"'.*?'", "'X'", msg_limpa)
        base = f"{tcode}|{msg_limpa}"
        contexto = ""
        
        if params and "Tipo de ordem" in params:
            contexto = f"|TIPO:{params['Tipo de ordem']}"

        base = f"{tcode}|{msg_limpa}{contexto}"
        return hashlib.md5(base.encode()).hexdigest()

    # LER DUMP
    def _read_dump(self, dump_path):

        if dump_path and os.path.exists(dump_path):

            try:
                with open(dump_path, "r", encoding="utf-8") as f:
                    return f.read()[:2000]

            except Exception:
                pass

        return "Sem dump."