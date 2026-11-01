"""Layout of the `categorised` output sheet."""

from __future__ import annotations

from dataclasses import astuple, dataclass
from enum import StrEnum


class Method(StrEnum):
    RULE = "rule"
    MODEL = "model"
    MANUAL = "manual"


CATEGORISED_COLUMNS: tuple[str, ...] = (
    "transactionId",
    "owner",
    "type",
    "category",
    "subCategory",
    "method",
    "confidence",
    "ruleId",
    "reviewed",
    "updatedAt",
    # Copied from `transactions` so a row (or a ruleId filter) shows what was categorised.
    # Added after the original columns so existing sheets are extended, not rebuilt.
    "bookingDate",
    "counterparty",
    "description",
    "amountEur",
)
CONTEXT_COLUMNS: tuple[str, ...] = ("bookingDate", "counterparty", "description", "amountEur")
KEY_COLUMN = CATEGORISED_COLUMNS.index("transactionId")
REVIEWED_COLUMN = CATEGORISED_COLUMNS.index("reviewed")


@dataclass(frozen=True)
class CategorisedRow:
    transactionId: str
    owner: str
    type: str
    category: str
    subCategory: str
    method: Method
    confidence: float | None  # None when not applicable (e.g. manual)
    ruleId: str
    reviewed: bool
    updatedAt: str  # ISO 8601, UTC
    bookingDate: str = ""
    counterparty: str = ""
    description: str = ""
    amountEur: float | str = ""

    def to_values(self) -> list:
        values = list(astuple(self))
        values[CATEGORISED_COLUMNS.index("method")] = str(self.method)
        if self.confidence is None:
            values[CATEGORISED_COLUMNS.index("confidence")] = ""
        return values


def transaction_context(txn: dict) -> dict:
    """The context columns (`bookingDate` … `amountEur`) of a transaction, as keyword arguments for `CategorisedRow`.
    `amountEur` falls back to `amount` when blank (same as `rules.amount_eur`)."""
    amount = txn.get("amountEur", "")
    if amount in ("", None):
        amount = txn.get("amount", "")
    return {
        "bookingDate": str(txn.get("bookingDate", "") or "")[:10],
        "counterparty": str(txn.get("counterparty", "") or "").strip(),
        "description": str(txn.get("description", "") or "").strip(),
        "amountEur": "" if amount is None else amount,
    }
