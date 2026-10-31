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
    issues_sheet: str = "categoriser_issues"
    rules_sheet: str = "categoriser_rules"
    taxonomy_sheet: str = "categoriser_taxonomy"
    owners_sheet: str = "account_owners"  # bank-sync's tab; read only (owner dropdown on rules)
    # Owner value used for shared accounts (the `owner` column on `transactions`).
    joint_owner: str = "common"
    # Legacy "Financial Planning" sheet: read only, by the mine/evaluate/seed commands.
    legacy_spreadsheet_id: str = ""

    def __post_init__(self) -> None:
        names = [n.strip() for n in (self.transactions_sheet, self.output_sheet, self.issues_sheet,
                                     self.rules_sheet, self.taxonomy_sheet)]
        if len(set(names)) != len(names):
            raise ValueError("TRANSACTIONS_SHEET, OUTPUT_SHEET, ISSUES_SHEET, RULES_SHEET and "
                             "TAXONOMY_SHEET must all differ.")

    @property
    def writable_sheets(self) -> frozenset[str]:
        """The only sheets this job may write to. Never the transactions sheet.

        The rules sheet is writable so the job can create it, fill `hits` and
        `lastMatched`, and append mined rules; the taxonomy sheet so it can be
        created and seeded once.
        """
        return frozenset({self.output_sheet, self.issues_sheet, self.rules_sheet, self.taxonomy_sheet})


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
        issues_sheet=os.environ.get("ISSUES_SHEET", "categoriser_issues"),
        rules_sheet=os.environ.get("RULES_SHEET", "categoriser_rules"),
        taxonomy_sheet=os.environ.get("TAXONOMY_SHEET", "categoriser_taxonomy"),
        joint_owner=os.environ.get("JOINT_OWNER", "common").strip().lower(),
        legacy_spreadsheet_id=os.environ.get("LEGACY_SPREADSHEET_ID", "").strip(),
    )
