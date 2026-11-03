"""Settings from the environment (`~/trade-republic/.env`, passed to the container by docker compose)."""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from zoneinfo import ZoneInfo


@dataclass(frozen=True)
class Settings:
    phone: str
    account: str
    tz: ZoneInfo
    data_dir: Path
    session_dir: Path
    output_name: str
    lookback_days: int
    drive_file_id: str
    credentials: Path

    @property
    def cookies_file(self) -> Path:
        return self.session_dir / "cookies.txt"

    @property
    def output_file(self) -> Path:
        return self.data_dir / self.output_name


def load(env: dict[str, str] | None = None) -> Settings:
    """Reads the settings; raises ValueError naming a missing or invalid one (never its value)."""
    e = os.environ if env is None else env
    home = Path(e.get("TR_HOME", "/app"))
    phone = e.get("TR_PHONE", "").strip()
    if not phone.startswith("+") or not phone[1:].isdigit():
        raise ValueError("TR_PHONE must be set in international format, e.g. +352...")
    try:
        lookback = int(e.get("TR_LOOKBACK_DAYS", "45"))
    except ValueError:
        raise ValueError("TR_LOOKBACK_DAYS must be a whole number of days") from None
    if lookback < 1:
        raise ValueError("TR_LOOKBACK_DAYS must be at least 1")
    return Settings(
        phone=phone,
        account=e.get("TR_ACCOUNT_NAME", "").strip() or "Trade Republic",
        tz=ZoneInfo(e.get("TR_TIME_ZONE", "").strip() or "Europe/Brussels"),
        data_dir=home / "data",
        session_dir=home / "session",
        output_name=e.get("TR_OUTPUT_FILE", "").strip() or "trade-republic.csv",
        lookback_days=lookback,
        drive_file_id=e.get("TR_DRIVE_FILE_ID", "").strip(),
        credentials=Path(e.get("GOOGLE_APPLICATION_CREDENTIALS", str(home / "credentials" / "service-account.json"))),
    )
