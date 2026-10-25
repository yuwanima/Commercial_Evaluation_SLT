import os
from dotenv import load_dotenv

BASE = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE, ".env"))

STORE = os.environ.get("STORE_DIR", os.path.join(BASE, "storage"))
SEED = os.path.join(BASE, "reference_seed")

def database_url():
    url = os.environ.get("DATABASE_URL", "").strip()
    if not url:
        os.makedirs(STORE, exist_ok=True)
        return "sqlite:///" + os.path.join(STORE, "checker.db")
    if url.startswith("postgres://"):
        url = "postgresql://" + url[len("postgres://"):]
    if url.startswith("postgresql://"):
        url = "postgresql+psycopg2://" + url[len("postgresql://"):]
    return url
