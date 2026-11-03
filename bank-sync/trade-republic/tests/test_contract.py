"""The CSV this job writes is the one Bank Sync reads: same columns, same order."""

import re
from pathlib import Path

from tr_export.csv_format import COLUMNS, SEPARATOR

JS = Path(__file__).resolve().parents[2] / "src" / "providers" / "csv" / "formats" / "trade-republic.js"


def test_columns_match_the_bank_sync_format():
    text = JS.read_text(encoding="utf-8")
    block = re.search(r"columns:\s*\[(.*?)\]", text, re.S).group(1)
    assert tuple(re.findall(r"'([^']+)'", block)) == COLUMNS
    assert re.search(r"separator:\s*'(.)'", text).group(1) == SEPARATOR
