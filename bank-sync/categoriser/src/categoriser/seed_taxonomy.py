"""Seed `categoriser_taxonomy` from the legacy sheet's `Dropdowns` tab: `python -m categoriser.seed_taxonomy`.

One-off. Does nothing if the taxonomy sheet already has rows, so the owner's
edits there are never overwritten. The legacy sheet is opened read-only.
"""

from __future__ import annotations

import logging

from .config import load_settings
from .legacy import DROPDOWNS_TAB
from .sheets import append_rows, ensure_taxonomy_sheet, open_readonly, open_spreadsheet, read_values
from .taxonomy import triples_from_dropdowns


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    log = logging.getLogger("categoriser")
    settings = load_settings()
    if not settings.legacy_spreadsheet_id:
        raise SystemExit("LEGACY_SPREADSHEET_ID is not set in .env.")
    ws = ensure_taxonomy_sheet(open_spreadsheet(settings), settings)
    if len(read_values(ws)) > 1:
        log.info("taxonomy sheet already has rows; nothing written")
        return
    dropdowns = read_values(open_readonly(settings, settings.legacy_spreadsheet_id).worksheet(DROPDOWNS_TAB))
    triples = triples_from_dropdowns(dropdowns)
    append_rows(ws, settings, [list(t) for t in triples])
    log.info("taxonomy seeded: %d rows", len(triples))


if __name__ == "__main__":
    main()
