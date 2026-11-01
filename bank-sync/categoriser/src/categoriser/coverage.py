"""After-write check that `categorised` covers `transactions` (pure, no I/O).

Counts what the run left uncategorised and reports, in the issues log:
- `transaction_id_not_unique`: a transactionId on several `transactions` rows (e.g. identical BIL CSV rows
  on two accounts get the same id). `categorised` is keyed by transactionId alone, so only the first row
  is categorised; bank-sync has to make the ids differ.
- `orphaned_categorised_row`: a `categorised` row whose transactionId is no longer on `transactions` (the
  row was removed, or bank-sync changed its id). Never deleted by the job, as it may hold your answer.
- `categorised_row_not_written`: a row this run labelled that is missing from `categorised` after the write.
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field

from .issues import Issue
from .merge import is_reviewed
from .schema import CATEGORISED_COLUMNS, KEY_COLUMN, REVIEWED_COLUMN

NOT_UNIQUE = "transaction_id_not_unique"
ORPHANED = "orphaned_categorised_row"
NOT_WRITTEN = "categorised_row_not_written"
LABEL_COLUMNS = tuple(CATEGORISED_COLUMNS.index(n) for n in ("type", "category", "subCategory"))


@dataclass
class Coverage:
    transactions: int = 0   # distinct transactionIds on `transactions`
    categorised: int = 0    # of those, with type, category and subCategory on `categorised`
    uncategorised: int = 0
    orphaned: int = 0
    not_written: int = 0
    not_unique: int = 0
    issues: list[Issue] = field(default_factory=list)


def check_coverage(transactions: list[dict], categorised: list[list], produced: set[str],
                   transactions_sheet: str, output_sheet: str) -> Coverage:
    """`categorised` is the output sheet as written, incl. header; `produced` the ids this run wrote."""
    source = f"categoriser.run: '{output_sheet}' vs '{transactions_sheet}' check"
    counts = Counter(str(t.get("transactionId", "")).strip() for t in transactions)
    counts.pop("", None)

    rows: dict[str, list[int]] = {}
    labelled: set[str] = set()
    reviewed: set[str] = set()
    for n, row in enumerate(categorised[1:], start=2):
        row = list(row) + [""] * (len(CATEGORISED_COLUMNS) - len(row))
        tid = str(row[KEY_COLUMN]).strip()
        if not tid:
            continue
        rows.setdefault(tid, []).append(n)
        if all(str(row[i]).strip() for i in LABEL_COLUMNS):
            labelled.add(tid)
        if is_reviewed(row[REVIEWED_COLUMN]):
            reviewed.add(tid)

    result = Coverage(transactions=len(counts))
    result.categorised = len(labelled & counts.keys())
    result.uncategorised = result.transactions - result.categorised

    for tid in sorted(t for t, c in counts.items() if c > 1):
        result.issues.append(Issue(
            NOT_UNIQUE, tid,
            f"transactionId is on {counts[tid]} rows of '{transactions_sheet}' (different accounts); only the "
            f"first is categorised. bank-sync has to give them different ids.", source))
    for tid in sorted(rows.keys() - counts.keys()):
        kept = "reviewed, so it holds your answer" if tid in reviewed else "not reviewed"
        result.issues.append(Issue(
            ORPHANED, tid,
            f"Row {', '.join(map(str, rows[tid]))} of '{output_sheet}' ({kept}) has no matching transaction on "
            f"'{transactions_sheet}' (removed, or its id changed). Left as is; delete the row if it is not needed.",
            source))
    for tid in sorted(produced - rows.keys()):
        result.issues.append(Issue(
            NOT_WRITTEN, tid,
            f"This run labelled the transaction but it is not on '{output_sheet}' after writing. "
            "The next run tries again; if it stays, the write is failing.", source))
    by_type = Counter(i.type for i in result.issues)
    result.not_unique, result.orphaned = by_type[NOT_UNIQUE], by_type[ORPHANED]
    result.not_written = by_type[NOT_WRITTEN]
    return result
