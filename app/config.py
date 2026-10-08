"""
Configuration module for the PO Translation System.

Handles environment variables and system configuration.
"""

import os
from pathlib import Path
from typing import List, Set

from dotenv import load_dotenv

# Base paths
BASE_DIR = Path(__file__).resolve().parent.parent

# Load environment variables from .env file
load_dotenv(BASE_DIR / ".env")


def _split_csv(value: str) -> List[str]:
    """Split a comma separated environment value into clean items."""
    return [item.strip() for item in value.split(",") if item.strip()]


# Application settings
class Settings:
    """Application settings loaded from environment."""

    def __init__(self) -> None:
        # Storage directories
        self.upload_dir: Path = Path(os.getenv("UPLOAD_DIR", str(BASE_DIR / "uploads")))
        self.output_dir: Path = Path(os.getenv("OUTPUT_DIR", str(BASE_DIR / "outputs")))

        # Translation provider configuration
        # google_free  - public Google Translate endpoint, no API key needed
        # google       - Google Cloud Translation API v2, needs TRANSLATION_API_KEY
        # mock         - offline fake translations, for tests and demos
        self.translation_provider: str = os.getenv("TRANSLATION_PROVIDER", "google_free").strip().lower()
        self.translation_api_key: str = os.getenv("TRANSLATION_API_KEY", "").strip()
        self.google_translate_api_endpoint: str = os.getenv(
            "GOOGLE_TRANSLATE_API_ENDPOINT",
            "https://translation.googleapis.com/language/translate/v2"
        )
        self.translation_timeout: int = int(os.getenv("TRANSLATION_TIMEOUT", "30"))

        # File handling settings
        self.max_file_size: int = int(os.getenv("MAX_FILE_SIZE", "10485760"))  # 10MB default
        self.allowed_extensions: Set[str] = {
            ext.lower() for ext in _split_csv(os.getenv("ALLOWED_EXTENSIONS", ".xlsx"))
        }

        # API / UI settings
        self.app_name: str = os.getenv("APP_NAME", "PO Translation System")
        self.cors_origins: List[str] = _split_csv(
            os.getenv("CORS_ORIGINS", "http://localhost:8501,http://127.0.0.1:8501")
        )

        # Ensure directories exist
        self.upload_dir.mkdir(parents=True, exist_ok=True)
        self.output_dir.mkdir(parents=True, exist_ok=True)


settings = Settings()
