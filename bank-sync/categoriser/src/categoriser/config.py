"""Settings loaded from the environment and the runtime folder's .env."""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

def categoriser_home() -> Path:
    """Runtime folder for .env, credentials/, data/, models/, logs/.

    Defaults to ~/categoriser: outside ~/Documents (macOS privacy blocks
    launchd jobs there) and outside iCloud and the public repo.
    """
    return Path(os.environ.get("CATEGORISER_HOME", Path.home() / "categoriser")).expanduser()


@dataclass(frozen=True)
class Settings:
    spreadsheet_id: str
    credentials_path: Path
    transactions_sheet: str = "transactions"
    output_sheet: str = "categorised"

    def __post_init__(self) -> None:
        if self.output_sheet.strip() == self.transactions_sheet.strip():
            raise ValueError("OUTPUT_SHEET must differ from TRANSACTIONS_SHEET.")


def load_settings() -> Settings:
    home = categoriser_home()
    load_dotenv(home / ".env")
    spreadsheet_id = os.environ.get("SPREADSHEET_ID", "").strip()
    if not spreadsheet_id:
        raise RuntimeError("SPREADSHEET_ID is not set (see .env.example).")

    creds = Path(os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "credentials/service-account.json"))
    if not creds.is_absolute():
        creds = home / creds
    if not creds.is_file():
        raise RuntimeError(f"Service-account key not found at {creds}.")

    return Settings(
        spreadsheet_id=spreadsheet_id,
        credentials_path=creds,
        transactions_sheet=os.environ.get("TRANSACTIONS_SHEET", "transactions"),
        output_sheet=os.environ.get("OUTPUT_SHEET", "categorised"),
    )
