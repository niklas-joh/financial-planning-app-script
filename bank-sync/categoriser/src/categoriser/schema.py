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
)
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
    confidence: float | str  # "" when not applicable (e.g. manual)
    ruleId: str
    reviewed: bool
    updatedAt: str  # ISO 8601, UTC

    def to_values(self) -> list:
        values = list(astuple(self))
        values[CATEGORISED_COLUMNS.index("method")] = str(self.method)
        return values
