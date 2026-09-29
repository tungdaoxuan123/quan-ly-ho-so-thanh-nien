"""Authentication and authorization configuration, secrets, and DB path resolution."""

from __future__ import annotations

import json
import os
import re
import secrets
import sys
from pathlib import Path
from typing import Optional, Union

from quan_ly_ho_so.config import BASE_DIR
from quan_ly_ho_so.state import settings_path


def load_env(env_path: Optional[Union[Path, str]] = None) -> None:
    """Load environment variables from the .env file at the main project directory.

    Tries python-dotenv first, with an automatic fallback parser if python-dotenv
    is not installed.
    """
    target = Path(env_path) if env_path is not None else BASE_DIR / ".env"
    if not target.is_file():
        return

    try:
        from dotenv import load_dotenv

        load_dotenv(dotenv_path=target)
        return
    except ImportError:
        pass

    # Built-in lightweight fallback parser for key=value .env files
    try:
        with open(target, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, val = line.split("=", 1)
                key = key.strip()
                val = val.strip()
                if (val.startswith('"') and val.endswith('"')) or (val.startswith("'") and val.endswith("'")):
                    val = val[1:-1]
                if key and key not in os.environ:
                    os.environ[key] = val
    except Exception:
        pass


# Ensure .env is loaded on module import
load_env()

USERNAME_REGEX = re.compile(r"^[a-zA-Z0-9_\-\.]{3,30}$")
DEFAULT_ADMIN_USERNAME = os.environ.get("DEFAULT_ADMIN_USERNAME", "admin")
DEFAULT_ADMIN_PASSWORD = os.environ.get("DEFAULT_ADMIN_PASSWORD", "admin123")
DEFAULT_USER_PASSWORD = os.environ.get("DEFAULT_USER_PASSWORD", "123456")
DEFAULT_SESSION_LIFETIME_DAYS = int(
    os.environ.get("PERMANENT_SESSION_LIFETIME_DAYS")
    or os.environ.get("SESSION_LIFETIME_DAYS")
    or "7"
)

# Test/runtime override for isolated testing
_auth_database_override: Optional[Path] = None


def set_auth_database_path(path: Optional[Union[Path, str]]) -> None:
    """Explicitly set an override path for the auth database (used in unit tests)."""
    global _auth_database_override
    _auth_database_override = Path(path).resolve() if path is not None else None


def auth_database_path() -> Path:
    """Return the SQLite database path for user authentication.

    Located at the main project directory, NOT based on Excel files or cache paths.
    Can be configured via AUTH_DB_PATH in the .env file.
    """
    if _auth_database_override is not None:
        return _auth_database_override

    env_path = os.environ.get("AUTH_DB_PATH", "").strip()
    if env_path:
        path = Path(env_path)
        if not path.is_absolute():
            path = BASE_DIR / path
        return path.resolve()

    return (BASE_DIR / "users.db").resolve()


def get_persistent_secret_key() -> str:
    """Return a stable secret key for Flask sessions.

    Checks SECRET_KEY or QUAN_LY_HO_SO_SECRET from environment/.env first.
    Falls back to settings.json or generates and persists a new random key.
    """
    env_secret = (
        os.environ.get("SECRET_KEY")
        or os.environ.get("QUAN_LY_HO_SO_SECRET")
        or os.environ.get("FLASK_SECRET_KEY")
    )
    if env_secret:
        return env_secret.strip()

    path = settings_path()
    try:
        if path.is_file():
            data = json.loads(path.read_text(encoding="utf-8"))
            if "flask_secret_key" in data and data["flask_secret_key"]:
                return data["flask_secret_key"]
    except Exception:
        pass

    new_secret = secrets.token_hex(32)
    try:
        data = {}
        if path.is_file():
            data = json.loads(path.read_text(encoding="utf-8"))
        data["flask_secret_key"] = new_secret
        path.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except Exception:
        pass
    return new_secret
