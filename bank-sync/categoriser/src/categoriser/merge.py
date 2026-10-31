"""Pure upsert planning for the `categorised` sheet (no I/O, easy to test).

Rules: an existing row whose `reviewed` cell is TRUE is never overwritten, and a
transactionId that appears on several rows is left alone on all of them (reported
in `UpsertPlan.duplicates` for the issues log) while the rest of the run continues.
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
    skipped_moved: list[str] = field(default_factory=list)  # row changed under us; next run retries
    duplicates: dict[str, list[int]] = field(default_factory=dict)  # transactionId -> sheet rows


def drop_newly_reviewed(plan: UpsertPlan, keys: list, reviewed: list) -> UpsertPlan:
    """Re-check planned updates against a fresh read of the key and reviewed columns.

    `keys` and `reviewed` are whole columns incl. header (index 0 = row 1). A planned
    update is dropped if its row is now reviewed, or no longer holds the same
    transactionId (the sheet was sorted or edited since the plan was made).
    """
    def cell(column: list, row_number: int) -> object:
        return column[row_number - 1] if len(column) >= row_number else ""

    kept: dict[int, list] = {}
    for row_number, values in plan.updates.items():
        key = values[KEY_COLUMN]
        if str(cell(keys, row_number)).strip() != key:
            plan.skipped_moved.append(key)
        elif is_reviewed(cell(reviewed, row_number)):
            plan.skipped_reviewed.append(key)
        else:
            kept[row_number] = values
    plan.updates = kept
    return plan


def plan_upsert(existing: list[list], incoming: list[CategorisedRow]) -> UpsertPlan:
    """`existing` is the full sheet incl. header (row 1), unformatted values."""
    if existing and [str(h) for h in existing[0][: len(CATEGORISED_COLUMNS)]] != list(CATEGORISED_COLUMNS):
        raise ValueError("Header row of the categorised sheet does not match the expected columns.")

    rows_by_key: dict[str, list[int]] = {}
    reviewed_by_row: dict[int, bool] = {}
    for offset, row in enumerate(existing[1:], start=2):
        key = str(row[KEY_COLUMN]).strip() if len(row) > KEY_COLUMN else ""
        if key:
            rows_by_key.setdefault(key, []).append(offset)
            reviewed_by_row[offset] = len(row) > REVIEWED_COLUMN and is_reviewed(row[REVIEWED_COLUMN])

    duplicates = {k: rows for k, rows in rows_by_key.items() if len(rows) > 1}
    index = {k: (rows[0], reviewed_by_row[rows[0]]) for k, rows in rows_by_key.items() if len(rows) == 1}

    plan = UpsertPlan(duplicates=duplicates)
    seen: set[str] = set()
    for item in incoming:
        key = item.transactionId.strip()
        if not key:
            raise ValueError("Every categorised row needs a transactionId.")
        if key in seen:
            raise ValueError(f"Duplicate transactionId in one run: {key}")
        seen.add(key)

        if key in duplicates:
            continue  # ambiguous: leave every copy alone, reported via plan.duplicates
        if key in index:
            row_number, reviewed = index[key]
            if reviewed:
                plan.skipped_reviewed.append(key)
            else:
                plan.updates[row_number] = item.to_values()
        else:
            plan.appends.append(item.to_values())
    return plan
