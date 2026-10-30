"""Pure upsert planning for the `categorised` sheet (no I/O, easy to test).

Rule: an existing row whose `reviewed` cell is TRUE is never overwritten.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .schema import CATEGORISED_COLUMNS, KEY_COLUMN, REVIEWED_COLUMN, CategorisedRow


def is_reviewed(value: object) -> bool:
    if isinstance(value, bool):
        return value
    return str(value).strip().upper() == "TRUE"


@dataclass
class UpsertPlan:
    updates: dict[int, list] = field(default_factory=dict)  # 1-based sheet row -> values
    appends: list[list] = field(default_factory=list)
    skipped_reviewed: list[str] = field(default_factory=list)


def plan_upsert(existing: list[list], incoming: list[CategorisedRow]) -> UpsertPlan:
    """`existing` is the full sheet incl. header (row 1), unformatted values."""
    if existing and [str(h) for h in existing[0][: len(CATEGORISED_COLUMNS)]] != list(CATEGORISED_COLUMNS):
        raise ValueError("Header row of the categorised sheet does not match the expected columns.")

    index: dict[str, tuple[int, bool]] = {}
    for offset, row in enumerate(existing[1:], start=2):
        key = str(row[KEY_COLUMN]).strip() if len(row) > KEY_COLUMN else ""
        if key:
            reviewed = len(row) > REVIEWED_COLUMN and is_reviewed(row[REVIEWED_COLUMN])
            index[key] = (offset, reviewed)

    plan = UpsertPlan()
    seen: set[str] = set()
    for item in incoming:
        key = item.transactionId.strip()
        if not key:
            raise ValueError("Every categorised row needs a transactionId.")
        if key in seen:
            raise ValueError(f"Duplicate transactionId in one run: {key}")
        seen.add(key)

        if key in index:
            row_number, reviewed = index[key]
            if reviewed:
                plan.skipped_reviewed.append(key)
            else:
                plan.updates[row_number] = item.to_values()
        else:
            plan.appends.append(item.to_values())
    return plan
