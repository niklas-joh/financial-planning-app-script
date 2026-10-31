"""Entry point: `python -m categoriser.run` (what Docker and launchd run).

I/O only for now: reads transactions, ensures the output sheet, writes
whatever the (not yet implemented) categoriser returns. Logs counts only,
never transaction contents.
"""

from __future__ import annotations

import logging
from datetime import UTC, datetime

from .config import load_settings
from .issues import DUPLICATE_ID, duplicate_issues
from .schema import CategorisedRow
from .sheets import (
    ensure_categorised_sheet,
    ensure_issues_sheet,
    open_spreadsheet,
    read_transactions,
    record_issues,
    upsert_categorised,
)

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

    now = datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
    issues = record_issues(
        ensure_issues_sheet(ss, settings), settings,
        duplicate_issues(plan.duplicates, settings.output_sheet), {DUPLICATE_ID}, now,
    )
    log.info(
        "read=%d updated=%d appended=%d skipped_reviewed=%d skipped_moved=%d skipped_duplicate=%d "
        "issues_opened=%d issues_reopened=%d issues_resolved=%d",
        len(transactions), len(plan.updates), len(plan.appends),
        len(plan.skipped_reviewed), len(plan.skipped_moved), len(plan.duplicates),
        issues.opened, issues.reopened, issues.resolved,
    )


if __name__ == "__main__":
    main()
