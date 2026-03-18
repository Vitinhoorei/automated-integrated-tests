from dataclasses import dataclass

@dataclass
class SapResult:
    status: str  # "PASS" ou "FAIL"
    source: str  # "STATUSBAR", "POPUP", "EXCEPTION", "OK"
    message: str
    evidence_path: str = ""