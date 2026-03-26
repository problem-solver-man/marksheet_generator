import sys
from pathlib import Path

# -------------------------
# SET RESOURCE PATHS
# -------------------------

if getattr(sys, "frozen", False):
    # Running as .exe
    base_dir = Path(sys.executable).parent
    temp_dir = Path(getattr(sys, "_MEIPASS"))
else:
    # Running as Python Script
    base_dir = Path(__file__).resolve().parent / "../"
    temp_dir = Path(__file__).resolve().parent / "../"