import json
import os
from threading import Lock

class ErrorRepository:
    def __init__(self, path="data/error/error_base.json"):
        self.path = path
        self.lock = Lock()
        self.db = self._load()

    def _load(self):
        if os.path.exists(self.path):
            with open(self.path, "r", encoding="utf-8") as f:
                return json.load(f)
        return {}

    def get_known_error(self, error_key):
        return self.db.get(error_key)

    def save_error(self, error_key, analysis):
        with self.lock:
            self.db[error_key] = analysis
            os.makedirs(os.path.dirname(self.path), exist_ok=True)
            with open(self.path, "w", encoding="utf-8") as f:
                json.dump(self.db, f, indent=2, ensure_ascii=False)