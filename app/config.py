"""Paths and settings for Decko Desktop. 100% local (spec D4)."""
import json
import os
from dataclasses import dataclass
from pathlib import Path

APP_DIR = Path(os.environ.get("APPDATA", str(Path.home()))) / "Decko"
ENGINE_DIR = APP_DIR / "engine"
DB_PATH = APP_DIR / "decko.db"
SETTINGS_PATH = APP_DIR / "settings.json"
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
    dock_mode: bool = True               # SP6: snap to active PowerPoint window
    decko_on_top: bool = False           # SP7: keep Decko above other windows
    resize_ppt_for_dock: bool = True     # SP7: shrink PowerPoint to free Decko width

    def validate(self) -> None:
        if self.provider not in VALID_PROVIDERS:
            raise ValueError(f"provider must be one of {VALID_PROVIDERS}")
        if self.provider == "generic" and not self.base_url:
            raise ValueError("generic provider requires base_url")


def ensure_app_dirs() -> None:
    APP_DIR.mkdir(parents=True, exist_ok=True)
    ENGINE_DIR.mkdir(parents=True, exist_ok=True)


def load_persisted() -> dict:
    """Settings + last-used deck persisted across restarts. The API key
    itself lives in the OS keyring, not here."""
    try:
        return json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_persisted(settings: "Settings", last_deck_path: str = "",
                    last_mode: str = "attach") -> None:
    ensure_app_dirs()
    SETTINGS_PATH.write_text(json.dumps({
        "provider": settings.provider,
        "model": settings.model,
        "base_url": settings.base_url,
        "dock_mode": settings.dock_mode,
        "decko_on_top": settings.decko_on_top,
        "resize_ppt_for_dock": settings.resize_ppt_for_dock,
        "last_deck_path": last_deck_path,
        "last_mode": last_mode,
    }, indent=2), encoding="utf-8")


def settings_from_persisted() -> "Settings":
    d = load_persisted()
    return Settings(
        provider=d.get("provider", "anthropic"),
        model=d.get("model", "claude-opus-4-7"),
        base_url=d.get("base_url", ""),
        dock_mode=bool(d.get("dock_mode", True)),
        decko_on_top=bool(d.get("decko_on_top", False)),
        resize_ppt_for_dock=bool(d.get("resize_ppt_for_dock", True)),
    )
