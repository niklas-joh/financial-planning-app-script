"""Entry point: `python -m categoriser.run` (what Docker and launchd run).

I/O only for now: reads transactions, ensures the output sheet, writes
whatever the (not yet implemented) categoriser returns. Logs counts only,
never transaction contents.
"""

from __future__ import annotations

import logging

from .config import load_settings
from .schema import CategorisedRow
from .sheets import ensure_categorised_sheet, open_spreadsheet, read_transactions, upsert_categorised

log = logging.getLogger("categoriser")


def categorise(transactions: list[dict]) -> list[CategorisedRow]:
    """Placeholder: rules + model come later."""
    return []


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    settings = load_settings()
    ss = open_spreadsheet(settings)
    ws = ensure_categorised_sheet(ss, settings)
    transactions = read_transactions(ss, settings)
    rows = categorise(transactions)
    plan = upsert_categorised(ws, settings, rows)
    log.info(
        "read=%d updated=%d appended=%d skipped_reviewed=%d",
        len(transactions), len(plan.updates), len(plan.appends), len(plan.skipped_reviewed),
    )


if __name__ == "__main__":
    main()
