"""Pure upsert planning for the `categorised` sheet (no I/O, easy to test).

Rules: an existing row whose `reviewed` cell is TRUE is never overwritten, and a
transactionId that appears on several rows is left alone on all of them (reported
in `UpsertPlan.duplicates` for the issues log) while the rest of the run continues.

A row whose values match what the run would write (ignoring `updatedAt`) is not
rewritten. With `clear_stale_at` set, an existing unreviewed row of one of the
`clear_methods` (by default `rule`) that this run no longer categorises (its rule was
switched off or changed, or the model is no longer confident) is cleared: only its
transactionId and `updatedAt` remain, so a wrong category does not linger.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .schema import CATEGORISED_COLUMNS, KEY_COLUMN, REVIEWED_COLUMN, CategorisedRow, Method


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
    cleared: list[str] = field(default_factory=list)  # stale rows emptied (a subset of updates)


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
    plan.cleared = [k for k in plan.cleared if any(v[KEY_COLUMN] == k for v in kept.values())]
    return plan


_UPDATED_AT = CATEGORISED_COLUMNS.index("updatedAt")


def _same_except_timestamp(row: list, values: list) -> bool:
    """An unchanged row is not rewritten, so `updatedAt` shows when it last changed."""
    def norm(v: object) -> str:
        if isinstance(v, bool):
            return str(v).upper()
        if isinstance(v, float) and v.is_integer():
            v = int(v)
        return str(v).strip()
    row = list(row) + [""] * (len(values) - len(row))
    return all(norm(a) == norm(b) for i, (a, b) in enumerate(zip(row, values)) if i != _UPDATED_AT)


def cleared_values(transaction_id: str, now: str) -> list:
    values = [""] * len(CATEGORISED_COLUMNS)
    values[KEY_COLUMN] = transaction_id
    values[REVIEWED_COLUMN] = False
    values[CATEGORISED_COLUMNS.index("updatedAt")] = now
    return values


def plan_upsert(existing: list[list], incoming: list[CategorisedRow], clear_stale_at: str | None = None,
                clear_methods: frozenset[Method] = frozenset({Method.RULE})) -> UpsertPlan:
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
            elif not _same_except_timestamp(existing[row_number - 1], item.to_values()):
                plan.updates[row_number] = item.to_values()
        else:
            plan.appends.append(item.to_values())

    if clear_stale_at is not None:
        method_col = CATEGORISED_COLUMNS.index("method")
        stale_methods = {str(m) for m in clear_methods}
        for key, (row_number, reviewed) in index.items():
            row = existing[row_number - 1]
            method = str(row[method_col]).strip() if len(row) > method_col else ""
            if key not in seen and not reviewed and method in stale_methods:
                plan.updates[row_number] = cleared_values(key, clear_stale_at)
                plan.cleared.append(key)
    return plan
