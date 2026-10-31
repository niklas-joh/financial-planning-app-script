"""Smoke test: `python -m categoriser.smoke` prints the transactions header row
and makes sure the `categorised` sheet exists. Prints no transaction data."""

from __future__ import annotations

from .config import load_settings
from .sheets import ensure_categorised_sheet, open_spreadsheet, read_transactions_header


def main() -> None:
    settings = load_settings()
    ss = open_spreadsheet(settings)
    print(f"Opened: {ss.title}")
    print(f"{settings.transactions_sheet} header: {read_transactions_header(ss, settings)}")
    ws = ensure_categorised_sheet(ss, settings)
    print(f"{ws.title} header: {ws.row_values(1)}")


if __name__ == "__main__":
    main()
