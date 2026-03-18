from dataclasses import dataclass
import os
from dotenv import load_dotenv

# Carrega o .env que está na RAIZ do projeto
load_dotenv(os.path.join(os.path.dirname(__file__), '..', '.env'))

IA_API_KEY = os.getenv("IA_API_KEY", "")
IA_BASE_URL = os.getenv("IA_BASE_URL", "")

@dataclass
class AppConfig:
    
    evidence_dir: str = os.getenv("EVIDENCE_DIR", "data/evidence")
    output_dir: str = os.getenv("OUTPUT_DIR", "data/output")