import os
from pathlib import Path

import dropbox
from dotenv import load_dotenv


def _load_env() -> Path | None:
    candidates = [
        Path.cwd() / ".env",
        Path(__file__).resolve().parent / ".env",
    ]
    for env_path in candidates:
        if env_path.exists():
            load_dotenv(dotenv_path=env_path, override=False)
            return env_path
    load_dotenv(override=False)
    return None


_LOADED_ENV_PATH = _load_env()


def get_dropbox_client() -> dropbox.Dropbox:
    app_key = os.getenv("DROPBOX_APP_KEY")
    app_secret = os.getenv("DROPBOX_APP_SECRET")
    refresh_token = os.getenv("DROPBOX_REFRESH_TOKEN")

    if not all([app_key, app_secret, refresh_token]):
        origin = str(_LOADED_ENV_PATH) if _LOADED_ENV_PATH else "nessun .env trovato"
        raise ValueError(
            "Credenziali Dropbox mancanti. "
            f"Controlla il file .env ({origin})."
        )

    return dropbox.Dropbox(
        app_key=app_key,
        app_secret=app_secret,
        oauth2_refresh_token=refresh_token,
        timeout=30,
    )

