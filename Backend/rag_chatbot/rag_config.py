
import os
from pathlib import Path

from dotenv import load_dotenv


# RAG chatbot folder
RAG_CHATBOT_DIR = Path(__file__).resolve().parent

# Load settings from the local .env file
ENV_FILE = RAG_CHATBOT_DIR / ".env"
load_dotenv(ENV_FILE)


# Google Vertex AI settings
GOOGLE_CLOUD_PROJECT = os.getenv("GOOGLE_CLOUD_PROJECT")
GOOGLE_CLOUD_LOCATION = os.getenv("GOOGLE_CLOUD_LOCATION")
GOOGLE_CLOUD_UNIVERSE_DOMAIN = os.getenv(
    "GOOGLE_CLOUD_UNIVERSE_DOMAIN",
    "googleapis.com"
)
GOOGLE_APPLICATION_CREDENTIALS = os.getenv(
    "GOOGLE_APPLICATION_CREDENTIALS"
)


# Embedding model settings
EMBEDDING_MODEL = os.getenv(
    "EMBEDDING_MODEL",
    "gemini-embedding-2"
)

EMBEDDING_DIMENSIONS = int(
    os.getenv("EMBEDDING_DIMENSIONS", "768")
)


def validate_vertex_settings():
    missing = []

    if not GOOGLE_CLOUD_PROJECT:
        missing.append("GOOGLE_CLOUD_PROJECT")

    if not GOOGLE_CLOUD_LOCATION:
        missing.append("GOOGLE_CLOUD_LOCATION")

    if not GOOGLE_APPLICATION_CREDENTIALS:
        missing.append("GOOGLE_APPLICATION_CREDENTIALS")
    elif not Path(GOOGLE_APPLICATION_CREDENTIALS).is_file():
        missing.append("Valid credentials file path")

    if missing:
        raise ValueError(
            "Vertex AI configuration is incomplete: "
            + ", ".join(missing)
        )

    return True


def validate_embedding_settings():
    if not EMBEDDING_MODEL:
        raise ValueError("EMBEDDING_MODEL is missing")

    if not 128 <= EMBEDDING_DIMENSIONS <= 3072:
        raise ValueError(
            "EMBEDDING_DIMENSIONS must be between 128 and 3072"
        )

    return True
