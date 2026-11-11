"""The CSV layout this job writes. Bank Sync reads it with `src/providers/csv/formats/trade-republic.js`;
a test checks both name the same columns in the same order, so change them together."""

from __future__ import annotations

import csv
import io
from decimal import Decimal
from typing import Iterable, Mapping

COLUMNS: tuple[str, ...] = (
    "TR event id",
    "Timestamp",
    "Date",
    "Account",
    "Event type",
    "Type",
    "Title",
    "Subtitle",
    "Amount",
    "Currency",
    "ISIN",
    "Shares",
    "Fees",
    "Taxes",
    "Classification",
)

SEPARATOR = ";"


def number(value: float | int | None) -> str:
    """Plain decimal text with a `.` point, no thousands separator and no exponent ('' for None).

    Bank Sync rejects anything else, so `1e-05` or `1,234.5` can never reach the sheet as a wrong value.
    """
    if value is None:
        return ""
    text = format(Decimal(str(value)).normalize(), "f")
    return "0" if text in ("-0", "") else text


def render(rows: Iterable[Mapping[str, str]]) -> bytes:
    """The whole file as UTF-8 bytes (no BOM, LF line ends), header first."""
    out = io.StringIO()
    writer = csv.DictWriter(out, fieldnames=COLUMNS, delimiter=SEPARATOR, lineterminator="\n", extrasaction="raise")
    writer.writeheader()
    for row in rows:
        writer.writerow(row)
    return out.getvalue().encode("utf-8")
