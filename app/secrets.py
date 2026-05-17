"""API key in Windows Credential Manager via keyring. Never in SQLite
or plaintext (spec §4)."""
import keyring

from app.config import KEYRING_SERVICE, KEYRING_USERNAME


def set_api_key(key: str) -> None:
    keyring.set_password(KEYRING_SERVICE, KEYRING_USERNAME, key)


def get_api_key() -> str | None:
    return keyring.get_password(KEYRING_SERVICE, KEYRING_USERNAME)


def has_api_key() -> bool:
    return bool(get_api_key())


def clear_api_key() -> None:
    try:
        keyring.delete_password(KEYRING_SERVICE, KEYRING_USERNAME)
    except keyring.errors.PasswordDeleteError:
        pass
