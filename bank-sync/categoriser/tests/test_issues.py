import pytest

from categoriser.issues import (
    COL,
    DUPLICATE_ID,
    ISSUE_COLUMNS,
    OPEN,
    RESOLVED,
    Issue,
    duplicate_issues,
    plan_issues,
)

HEADER = list(ISSUE_COLUMNS)
T1, T2 = "2026-10-31T06:15:00Z", "2026-11-01T06:15:00Z"


def dup(tid="a", rows=(2, 4)):
    return duplicate_issues({tid: list(rows)}, "categorised")[0]


def logged(issue, status=OPEN, first=T1, last=T1, resolved=""):
    return [issue.key, status, issue.type, issue.transactionId, issue.detail, issue.source, first, last, resolved]


def test_new_issue_is_appended_once_as_open():
    plan = plan_issues([HEADER], [dup()], {DUPLICATE_ID}, T1)
    assert len(plan.appends) == 1 and plan.opened == 1
    row = plan.appends[0]
    assert row[COL["issueKey"]] == "duplicate_transaction_id:a"
    assert row[COL["status"]] == OPEN
    assert row[COL["firstSeen"]] == row[COL["lastSeen"]] == T1
    assert "rows 2, 4" in row[COL["detail"]]


def test_issue_seen_again_is_refreshed_not_repeated():
    plan = plan_issues([HEADER, logged(dup())], [dup()], {DUPLICATE_ID}, T2)
    assert plan.appends == [] and plan.opened == 0
    row = plan.updates[2]
    assert row[COL["firstSeen"]] == T1 and row[COL["lastSeen"]] == T2


def test_detail_follows_the_current_rows():
    plan = plan_issues([HEADER, logged(dup(rows=(2, 4)))], [dup(rows=(2, 9))], {DUPLICATE_ID}, T2)
    assert "rows 2, 9" in plan.updates[2][COL["detail"]]


def test_issue_gone_is_resolved():
    plan = plan_issues([HEADER, logged(dup())], [], {DUPLICATE_ID}, T2)
    assert plan.resolved == 1
    assert plan.updates[2][COL["status"]] == RESOLVED
    assert plan.updates[2][COL["resolvedAt"]] == T2


def test_resolved_issue_is_left_alone_while_absent():
    plan = plan_issues([HEADER, logged(dup(), status=RESOLVED, resolved=T1)], [], {DUPLICATE_ID}, T2)
    assert plan.updates == {} and plan.appends == []


def test_resolved_issue_that_returns_is_reopened_in_place():
    plan = plan_issues([HEADER, logged(dup(), status=RESOLVED, resolved=T1)], [dup()], {DUPLICATE_ID}, T2)
    assert plan.reopened == 1 and plan.appends == []
    row = plan.updates[2]
    assert row[COL["status"]] == OPEN and row[COL["resolvedAt"]] == "" and row[COL["firstSeen"]] == T1


def test_unchecked_issue_types_are_not_resolved():
    other = Issue("some_future_check", "x", "detail", "source")
    plan = plan_issues([HEADER, logged(other)], [], {DUPLICATE_ID}, T2)
    assert plan.updates == {}


def test_wrong_header_is_rejected():
    with pytest.raises(ValueError):
        plan_issues([["when", "what"]], [dup()], {DUPLICATE_ID}, T1)
