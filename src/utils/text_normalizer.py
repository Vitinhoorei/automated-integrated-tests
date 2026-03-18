import unicodedata
import re

def normalize_text(text: str) -> str:
    if not text:
        return ""

    text = "".join(
        ch for ch in unicodedata.normalize("NFD", text)
        if unicodedata.category(ch) != "Mn"
    )
    
    text = text.lower().strip()
    text = re.sub(r"[^\w\s]", " ", text)
    return " ".join(text.split())

def sanitize_filename(text: str) -> str:
    if not text:
        return "UNK"
    text = re.sub(r"[^\w\-]+", "_", text.strip())
    return text[:60]