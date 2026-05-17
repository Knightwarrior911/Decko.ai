"""Paths and settings for Decko Desktop. 100% local (spec D4)."""
import os
from dataclasses import dataclass
from pathlib import Path

APP_DIR = Path(os.environ.get("APPDATA", str(Path.home()))) / "Decko"
ENGINE_DIR = APP_DIR / "engine"
DB_PATH = APP_DIR / "decko.db"
CARRIER_NAME = "PPT_AI_Editor.pptm"
INSTALLED_CARRIER = ENGINE_DIR / CARRIER_NAME

KEYRING_SERVICE = "DeckoDesktop"
KEYRING_USERNAME = "llm_api_key"

VALID_PROVIDERS = ("anthropic", "openai", "generic")


@dataclass
class Settings:
    provider: str = "anthropic"          # one of VALID_PROVIDERS
    model: str = "claude-opus-4-7"
    base_url: str = ""                   # required only when provider == "generic"

    def validate(self) -> None:
        if self.provider not in VALID_PROVIDERS:
            raise ValueError(f"provider must be one of {VALID_PROVIDERS}")
        if self.provider == "generic" and not self.base_url:
            raise ValueError("generic provider requires base_url")


def ensure_app_dirs() -> None:
    APP_DIR.mkdir(parents=True, exist_ok=True)
    ENGINE_DIR.mkdir(parents=True, exist_ok=True)
