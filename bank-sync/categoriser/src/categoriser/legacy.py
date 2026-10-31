"""The legacy "Financial Planning" sheet: hand-labelled history used to mine and test rules.

Only read (with a read-only token, see `sheets.open_readonly`); never written.
Columns are found by header name; where a name repeats, the first one counts.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, timedelta

from .rules import normalise
from .transfers import MAX_DAYS, _cents, _day

TRANSACTIONS_TAB = "All Transactions"
DROPDOWNS_TAB = "Dropdowns"
_SHEETS_EPOCH = date(1899, 12, 30)


@dataclass(frozen=True)
class LegacyRow:
    day: date
    amount: float
    source: str
    type: str
    category: str
    subCategory: str
    counterparty: str
    description: str

    @property
    def label(self) -> tuple[str, str, str]:
        return (self.type, self.category, self.subCategory)

    def as_transaction(self) -> dict:
        """The fields rules look at, in `transactions` shape (no account: the legacy
        accounts are named differently)."""
        return {"counterparty": self.counterparty, "description": self.description,
                "reference": "", "amount": self.amount, "bookingDate": self.day.isoformat()}


def _date(value: object) -> date | None:
    if isinstance(value, (int, float)) and not isinstance(value, bool) and value > 0:
        return _SHEETS_EPOCH + timedelta(days=int(value))
    return _day(value)


def parse_legacy(values: list[list]) -> list[LegacyRow]:
    """Labelled rows of the legacy transactions tab (header first). Unlabelled rows are dropped."""
    if not values:
        return []
    header = [str(h).strip() for h in values[0]]
    col = {}
    for i, name in enumerate(header):
        col.setdefault(name, i)
    needed = ("Source", "Type", "Category", "Sub-Category", "Date", "Amount", "Beneficiary Name", "Operation")
    missing = [n for n in needed if n not in col]
    if missing:
        raise ValueError(f"Legacy sheet is missing columns: {', '.join(missing)}")

    def cell(row: list, name: str) -> object:
        i = col.get(name)
        return row[i] if i is not None and len(row) > i else ""

    out = []
    for row in values[1:]:
        label = [str(cell(row, n)).strip() for n in ("Type", "Category", "Sub-Category")]
        day = _date(cell(row, "Date"))
        try:
            amount = float(cell(row, "Amount"))
        except (TypeError, ValueError):
            continue
        if not all(label) or day is None:
            continue
        description = " ".join(str(cell(row, n)).strip() for n in ("Operation", "Communication")).strip()
        out.append(LegacyRow(day, amount, str(cell(row, "Source")).strip(), *label,
                             counterparty=str(cell(row, "Beneficiary Name")).strip(), description=description))
    return out


def join_legacy(legacy: list[LegacyRow], transactions: list[dict]) -> dict[str, LegacyRow]:
    """transactionId -> legacy row for the same booking: same amount to the cent (in the
    account's currency), at most `MAX_DAYS` apart, one to one; closest dates first,
    then the most words in common (two same-amount bookings on one day)."""
    by_cents: dict[int, list[tuple[str, date, set[str]]]] = {}
    for t in transactions:
        tid, cents, day = str(t.get("transactionId", "")).strip(), _cents(t.get("amount")), _day(t.get("bookingDate"))
        if tid and cents is not None and day:
            words = set(normalise(f"{t.get('counterparty', '')} {t.get('description', '')}").split())
            by_cents.setdefault(cents, []).append((tid, day, words))
    candidates = []
    for j, row in enumerate(legacy):
        words = set(normalise(f"{row.counterparty} {row.description}").split())
        for tid, day, txn_words in by_cents.get(_cents(row.amount), []):
            gap = abs((day - row.day).days)
            if gap <= MAX_DAYS:
                candidates.append((gap, -len(words & txn_words), j, tid))
    candidates.sort()
    joined: dict[str, LegacyRow] = {}
    used: set[int] = set()
    for _gap, _overlap, j, tid in candidates:
        if j in used or tid in joined:
            continue
        used.add(j)
        joined[tid] = legacy[j]
    return joined
