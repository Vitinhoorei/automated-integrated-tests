import os
import re
import json
import hashlib
import requests
import config

from param_enricher import enrich_params

class AITestIntegrator:
    """
    Responsabilidades:
    - Chamar IA de forma controlada
    - NÃO criar parâmetros livremente
    - Analisar erros SAP com base em:
        - Mensagem SAP
        - Dump (se existir)
        - Histórico já executado
    - Gerar sugestão + justificativa + confiança
    """

    def __init__(self):
        self.api_url = getattr(config, "IA_BASE_URL", os.getenv("IA_BASE_URL", ""))
        self.api_key = getattr(config, "IA_API_KEY", os.getenv("IA_API_KEY", ""))

        # Base viva de erros já vistos na execução atual
        self.error_memory: dict[str, dict] = {}

        # IDs gerados em execuções integradas
        self.shared_context: dict[str, str] = {}

    # IA – chamada controlada
    def _chamar_ia(self, prompt: str, json_mode: bool = False) -> str:
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
                return "{}" if json_mode else "Erro ao chamar IA."

            text = resp.json().get("response", "")
            if json_mode:
                match = re.search(r"\{.*\}", text, re.DOTALL)
                return match.group(0) if match else "{}"

            return text

        except Exception as e:
            return "{}" if json_mode else f"Falha IA: {e}"

    # Preparação de parâmetros (SEM alucinação e COM Teste Integrado)
    def preparar_parametros(
        self,
        tcode: str,
        explanation: str,
        raw_params: str
    ) -> dict[str, str]:
        
        params_basicos = enrich_params(tcode, explanation, raw_params)

        prompt = f"""
                    Você é um especialista SAP PM.
                    Sua tarefa é mesclar os parâmetros básicos já encontrados com o CONTEXTO de testes integrados.

                    TCODE: {tcode}
                    Descrição: {explanation}
                    Parâmetros Básicos já extraídos: {json.dumps(params_basicos, ensure_ascii=False)}
                    
                    CONTEXTO GERADO ANTERIORMENTE (Use esses IDs se a descrição pedir):
                    {json.dumps(self.shared_context, ensure_ascii=False)}

                    REGRAS:
                    - Retorne SOMENTE JSON válido {{ "Campo": "Valor" }}
                    - Não invente campos que não estejam na descrição ou no contexto.
                    - Não use camelCase.
                    """
        resposta = self._chamar_ia(prompt, json_mode=True)

        try:
            data = json.loads(resposta)
            return {k: str(v) for k, v in data.items() if isinstance(v, (str, int))}
        except Exception:
            return params_basicos or {} # Fallback seguro

    # Captura de IDs gerados (execuções integradas)
    def extrair_id_integrado(self, tcode: str, status_message: str) -> None:
        if not status_message:
            return

        prompt = f"""
                    A transação {tcode} retornou:
                    "{status_message}"

                    Extraia a entidade e o ID gerado.
                    Retorne JSON:
                    {{ "entidade": "NOME", "id": "NUMERO" }}
                    """
        resposta = self._chamar_ia(prompt, json_mode=True)

        try:
            data = json.loads(resposta)
            if data.get("entidade") and data.get("id"):
                self.shared_context[data["entidade"]] = str(data["id"])
        except Exception:
            pass

    # Análise de erro SAP (núcleo do seu projeto)
    def analisar_erro_sap(
        self,
        tcode: str,
        status_message: str,
        dump_path: str | None
    ) -> dict:
        """
        Retorna:
        - causa_raiz
        - sugestao_correcao
        - justificativa
        - confianca (0–100)
        """

        if not status_message:
            return {}

        normalized = self._normalize_error(tcode, status_message)

        # 1️⃣ Erro já conhecido
        if normalized in self.error_memory:
            prev = self.error_memory[normalized]
            return {
                "causa_raiz": prev["causa_raiz"],
                "sugestao_correcao": prev["sugestao_correcao"],
                "justificativa": "Erro idêntico já ocorrido anteriormente no mesmo fluxo.",
                "confianca": 90
            }

        # 2️⃣ Erro novo → análise profunda
        dump_text = self._read_dump(dump_path)

        prompt = f"""
                    Analise a falha SAP abaixo.

                    TCODE: {tcode}
                    Mensagem SAP: "{status_message}"
                    Dump (se houver):
                    {dump_text}

                    Considere:
                    - campos obrigatórios
                    - parâmetros ausentes ou inválidos
                    - etapa do fluxo

                    Retorne APENAS JSON:
                    {{
                    "causa_raiz": "até 100 caracteres",
                    "sugestao_correcao": "até 150 caracteres"
                    }}
                    """
        resposta = self._chamar_ia(prompt, json_mode=True)

        try:
            data = json.loads(resposta)
        except Exception:
            data = {
                "causa_raiz": "Erro não identificado",
                "sugestao_correcao": "Verificar mensagem SAP e evidência."
            }

        # 3️⃣ Armazena para próximas linhas
        self.error_memory[normalized] = data

        return {
            **data,
            "justificativa": "Primeira ocorrência deste erro. Análise baseada na mensagem SAP e no dump.",
            "confianca": 65
        }

    # Helpers
    def _normalize_error(self, tcode: str, message: str) -> str:
        base = f"{tcode}|{message.lower().strip()}"
        return hashlib.md5(base.encode("utf-8")).hexdigest()

    def _read_dump(self, dump_path: str | None) -> str:
        if dump_path and os.path.exists(dump_path):
            try:
                with open(dump_path, "r", encoding="utf-8") as f:
                    return f.read()[:2000]
            except Exception:
                pass
        return "Sem dump."