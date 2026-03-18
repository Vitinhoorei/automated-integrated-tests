import json
import re
from src.config import IA_API_KEY, IA_BASE_URL
import requests

class AITestIntegrator:
    def __init__(self):
        self.api_key = IA_API_KEY
        self.url = IA_BASE_URL
        self.shared_context = {}

    def _call_ia(self, prompt):
        headers = {"apiKey": self.api_key, "Content-Type": "application/json"}
        payload = {"question": prompt, "model": "gpt-4o-mini", "temperature": 0}
        try:
            r = requests.post(self.url, headers=headers, json=payload, timeout=60)
            return r.json().get("response", "")
        except: return "{}"

    def extrair_id_integrado(self, tcode, message):
        """Usa a IA para ler a barra de status e salvar IDs (ex: Ordem 123)."""
        prompt = f"Extraia o ID e o tipo de objeto da mensagem SAP: '{message}'. Retorne JSON: {{'tipo': 'Ordem', 'id': '123'}}"
        res = self._call_ia(prompt)
        try:
            data = json.loads(re.search(r"\{.*\}", res).group(0))
            if data.get("id"):
                self.shared_context[data["tipo"]] = data["id"]
        except: pass

    def analisar_erro_sap(self, tcode, msg, params):
        prompt = f"""
                    Erro no SAP. TCode: {tcode}. Mensagem: {msg}.
                    Parâmetros atuais: {params}.
                    Sugira uma correção no formato Campo=Valor se possível.
                    Retorne JSON: {{"causa": "...", "sugestao": "Campo=Valor", "confianca": 80}}
                """
                
        res = self._call_ia(prompt)
        try:
            return json.loads(re.search(r"\{.*\}", res).group(0))
        except:
            return {"confianca": 0}