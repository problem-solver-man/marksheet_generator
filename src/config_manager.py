from utils import base_dir
from platformdirs import user_config_dir
from pathlib import Path
import json

# Application info
APP_NAME = "MarksheetGenerator"
APP_AUTHOR = "Vishesh"
APP_VERSION = "1.0"


# -------------------------
# CONFIG
# -------------------------

# Create folder/file to save user preferences in config.json
config_dir = Path(user_config_dir(APP_NAME, APP_AUTHOR, APP_VERSION))
config_dir.mkdir(parents=True, exist_ok=True)

CONFIG_FILE = config_dir / "config.json"


# Set default output location
default_output = base_dir / "Output"
default_output.mkdir(parents=True, exist_ok=True)


# Function to load last user preferences
def load_config():
    if CONFIG_FILE.exists():
        with CONFIG_FILE.open() as f:
            return json.load(f)

    return {
        "last_excel_file": str(base_dir / "sample_data/Sample.xlsx"),
        "last_output_folder": str(default_output),
        "last_class": "All Class",
        "marksheet_order": "Order by Rank",
        "theme": "light"
    }


# Function to save user preferences
def save_config():
    with CONFIG_FILE.open("w") as f:
        json.dump(config, f, indent=4)


# Get last user preferences if available or set default
config = load_config()