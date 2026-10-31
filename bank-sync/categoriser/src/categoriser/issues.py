"""Issues log: problems the job found and skipped, one row per issue.

An issue is identified by `issueKey` (`<type>:<transactionId>`). Each run:
- a new issue is appended as `open`;
- an issue still present only gets `lastSeen` (and `detail`) refreshed, so the
  log does not grow on every run;
- an open issue no longer found is marked `resolved` with `resolvedAt`;
- a resolved issue that comes back is reopened in place.
Only issue types the run actually checked are resolved, so adding a new check
never closes another check's issues.
"""

from __future__ import annotations

from dataclasses import dataclass, field

ISSUE_COLUMNS: tuple[str, ...] = (
    "issueKey",
    "status",
    "type",
    "transactionId",
    "detail",
    "source",
    "firstSeen",
    "lastSeen",
    "resolvedAt",
)
COL = {name: i for i, name in enumerate(ISSUE_COLUMNS)}

OPEN, RESOLVED = "open", "resolved"
DUPLICATE_ID = "duplicate_transaction_id"


@dataclass(frozen=True)
class Issue:
    type: str
    transactionId: str
    detail: str
    source: str
    ref: str = ""  # what the issue is about when it is not a transaction, e.g. a ruleId

    @property
    def key(self) -> str:
        return f"{self.type}:{self.ref or self.transactionId}"


@dataclass
class IssuesPlan:
    updates: dict[int, list] = field(default_factory=dict)  # 1-based sheet row -> full row
    appends: list[list] = field(default_factory=list)
    opened: int = 0
    reopened: int = 0
    resolved: int = 0


def duplicate_issues(duplicates: dict[str, list[int]], sheet: str) -> list[Issue]:
    return [
        Issue(
            type=DUPLICATE_ID,
            transactionId=tid,
            detail=f"transactionId is on rows {', '.join(map(str, rows))} of '{sheet}'; "
                   "all copies left unchanged. Delete the extra rows to resolve.",
            source=f"categoriser.run: '{sheet}' sheet check",
        )
        for tid, rows in sorted(duplicates.items())
    ]


def plan_issues(existing: list[list], current: list[Issue], checked_types: set[str], now: str) -> IssuesPlan:
    """`existing` is the full issues sheet incl. header (row 1)."""
    if existing and [str(h) for h in existing[0][: len(ISSUE_COLUMNS)]] != list(ISSUE_COLUMNS):
        raise ValueError("Header row of the issues sheet does not match the expected columns.")

    def padded(row: list) -> list:
        return list(row[: len(ISSUE_COLUMNS)]) + [""] * (len(ISSUE_COLUMNS) - len(row))

    rows_by_key: dict[str, tuple[int, list]] = {}
    for offset, row in enumerate(existing[1:], start=2):
        row = padded(row)
        key = str(row[COL["issueKey"]]).strip()
        if key:
            rows_by_key[key] = (offset, row)  # last row wins if the log was edited by hand

    plan = IssuesPlan()
    current_keys = set()
    for issue in current:
        current_keys.add(issue.key)
        if issue.key not in rows_by_key:
            plan.appends.append([issue.key, OPEN, issue.type, issue.transactionId, issue.detail,
                                 issue.source, now, now, ""])
            plan.opened += 1
            continue
        row_number, row = rows_by_key[issue.key]
        if row[COL["status"]] != OPEN:
            plan.reopened += 1
        row[COL["status"]] = OPEN
        row[COL["detail"]] = issue.detail
        row[COL["lastSeen"]] = now
        row[COL["resolvedAt"]] = ""
        plan.updates[row_number] = row

    for key, (row_number, row) in rows_by_key.items():
        if key in current_keys or row[COL["status"]] != OPEN or row[COL["type"]] not in checked_types:
            continue
        row[COL["status"]] = RESOLVED
        row[COL["resolvedAt"]] = now
        plan.updates[row_number] = row
        plan.resolved += 1
    return plan
