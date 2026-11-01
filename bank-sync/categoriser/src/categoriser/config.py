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
    review_sheet: str = "categoriser_review"
    owners_sheet: str = "account_owners"  # bank-sync's tab; read only (owner dropdown on rules)
    # Owner value used for shared accounts (the `owner` column on `transactions`).
    joint_owner: str = "common"
    # Legacy "Financial Planning" sheet: read only, by the mine/evaluate/seed/train commands.
    legacy_spreadsheet_id: str = ""
    # The model labels a row only when its probability is at least this (0-1).
    model_min_confidence: float = 0.9
    # Review inbox: rows per run, of which this many are spot checks of model rows.
    review_size: int = 25
    review_spot_checks: int = 3

    def __post_init__(self) -> None:
        names = [n.strip() for n in (self.transactions_sheet, self.output_sheet, self.issues_sheet,
                                     self.rules_sheet, self.taxonomy_sheet, self.review_sheet)]
        if len(set(names)) != len(names):
            raise ValueError("TRANSACTIONS_SHEET, OUTPUT_SHEET, ISSUES_SHEET, RULES_SHEET, TAXONOMY_SHEET and "
                             "REVIEW_SHEET must all differ.")
        if not 0 < self.model_min_confidence <= 1:
            raise ValueError("MODEL_MIN_CONFIDENCE must be above 0 and at most 1.")
        if self.review_size < 1 or not 0 <= self.review_spot_checks <= self.review_size:
            raise ValueError("REVIEW_SIZE must be at least 1 and REVIEW_SPOT_CHECKS between 0 and REVIEW_SIZE.")

    @property
    def writable_sheets(self) -> frozenset[str]:
        """The only sheets this job may write to. Never the transactions sheet.

        The rules sheet is writable so the job can create it, add missing job
        columns, fill `hits`, `wins`, `lastMatched` and `sample`, add a filter and
        append mined rules and rules made in the review inbox; the taxonomy sheet so
        it can be created and seeded once; the review sheet is the inbox the job
        fills each run.
        """
        return frozenset({self.output_sheet, self.issues_sheet, self.rules_sheet, self.taxonomy_sheet,
                          self.review_sheet})


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
        review_sheet=os.environ.get("REVIEW_SHEET", "categoriser_review"),
        joint_owner=os.environ.get("JOINT_OWNER", "common").strip().lower(),
        legacy_spreadsheet_id=os.environ.get("LEGACY_SPREADSHEET_ID", "").strip(),
        model_min_confidence=_float("MODEL_MIN_CONFIDENCE", 0.9),
        review_size=_int("REVIEW_SIZE", 25),
        review_spot_checks=_int("REVIEW_SPOT_CHECKS", 3),
    )


def _float(name: str, default: float) -> float:
    value = os.environ.get(name, "").strip()
    try:
        return float(value) if value else default
    except ValueError:
        raise ValueError(f"{name} must be a number, e.g. {default}.") from None


def _int(name: str, default: int) -> int:
    value = os.environ.get(name, "").strip()
    try:
        return int(value) if value else default
    except ValueError:
        raise ValueError(f"{name} must be a whole number, e.g. {default}.") from None


def models_dir() -> Path:
    """Where `categoriser.train` saves the model and the daily run loads it."""
    return categoriser_home() / "models"


def data_dir() -> Path:
    """Runtime data (holdout reports, review history); never the repo."""
    return categoriser_home() / "data"
