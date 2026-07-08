from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent # where python file lives (cd ..)
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"

FAM_DIR = INPUT_DIR / "FAM"
SAM_DIR = INPUT_DIR / "SAM"
MOHAWK_DIR = INPUT_DIR / "MOHAWK"
LOCAL_DIR = INPUT_DIR / "LOCAL"
