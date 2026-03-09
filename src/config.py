from dataclasses import dataclass
import os
from dotenv import load_dotenv

IA_API_KEY = "FDiPWKBl7PT5ZargE8Gj7C05qnOoeQGB"
IA_BASE_URL = "https://apihub.weg.net/external/fb8df9df-d49b-42b3-8da0-794e175cab44/v1/chat_genai"

# carrega .env (se existir)
load_dotenv()

@dataclass
class AppConfig:
    """
    Config do projeto.
    Você pode trocar paths sem mexer no código, usando .env.
    """
    evidence_dir: str = os.getenv("EVIDENCE_DIR", "C:/Users/veleoterio/Downloads/testes-automação/evidence")
    output_dir: str = os.getenv("OUTPUT_DIR", "data/output")
    