import json
import os
from threading import Lock


class ErrorRepository:

    def __init__(self, path="data/error/error_base.json"):
        self.path = path
        self.lock = Lock()
        self.errors = {}

        self._load()

    def _load(self):

        if os.path.exists(self.path):

            try:
                with open(self.path, "r", encoding="utf-8") as f:
                    self.errors = json.load(f)
            except Exception:
                self.errors = {}

    def _save(self):

        os.makedirs(os.path.dirname(self.path), exist_ok=True)

        with self.lock:
            with open(self.path, "w", encoding="utf-8") as f:
                json.dump(self.errors, f, indent=2, ensure_ascii=False)

    def get(self, key):

        return self.errors.get(key)

    def save(self, key, data):

        self.errors[key] = data
        self._save()