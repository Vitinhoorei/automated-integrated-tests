from dataclasses import dataclass
import os
from dotenv import load_dotenv

# carrega .env antes de ler as variáveis
load_dotenv()

IA_API_KEY = os.getenv("IA_API_KEY", "")
IA_BASE_URL = os.getenv(
    "IA_BASE_URL",
    "https://apihub.weg.net/external/fb8df9df-d49b-42b3-8da0-794e175cab44/v1/chat_genai",
)

@dataclass
class AppConfig:
    """
    Config do projeto.
    Você pode trocar paths sem mexer no código, usando .env.
    """
    evidence_dir: str = os.getenv("EVIDENCE_DIR", "data/evidence")
    output_dir: str = os.getenv("OUTPUT_DIR", "data/output")