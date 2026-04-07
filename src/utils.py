

# -------------------------
# SET RESOURCE PATHS
# -------------------------
import sys
from pathlib import Path

if getattr(sys, "frozen", False):
    # Running as .exe (PyInstaller)
    base_dir = Path(sys.executable).parent
    temp_dir = Path(getattr(sys, "_MEIPASS"))
else:
    # Running as Python Script
    base_dir = Path(__file__).resolve().parents[1]
    temp_dir = Path(__file__).resolve().parents[1]



# -------------------------
# LOG MANAGER
# -------------------------
import logging

class CTkinterHandler(logging.Handler):
    def __init__(self, log_box) -> None:
        super().__init__()
        self.log_box = log_box

    def emit(self, record: logging.LogRecord) -> None:
        try:
            if not self.log_box.winfo_exists():
                return
            
            msg = self.format(record)

            self.log_box.configure(state="normal")
            self.log_box.insert("end", msg + "\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
            self.log_box.update_idletasks()
        
        except Exception:
            pass



# -------------------------
# LANGUAGE MANAGER
# -------------------------

import json
from pathlib import Path

class LanguageManager:
    def __init__(self, default_lang):
        self.lang = default_lang
        self.translations = {}
        self.load_language()

    def load_language(self):
        path = Path(f"{temp_dir}/languages/{self.lang}.json")
        with open(path, "r", encoding="utf-8") as f:
            self.translations = json.load(f)

    def set_language(self, lang):
        self.lang = lang
        self.load_language()

    def t(self, key):
        return self.translations.get(key, key)