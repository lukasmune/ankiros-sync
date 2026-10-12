from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
EXPORTS_DIR = BASE_DIR / "exports"
DATABASE_PATH = DATA_DIR / "ankiros.db"

API_BASE_URL = "https://web.ankirosforms.com/"
EXPO_ID = "a069ef48-11a1-405b-5070-08de2b499b9f"
EDITION = 2026
REQUEST_TIMEOUT = 30
REQUEST_RETRIES = 4
REQUEST_DELAY = 0.5
