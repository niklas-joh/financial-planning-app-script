from pathlib import Path

import pytest

from categoriser.config import Settings
from categoriser.issues import COL, DUPLICATE_ID, ISSUE_COLUMNS, duplicate_issues
from categoriser.schema import CATEGORISED_COLUMNS, REVIEWED_COLUMN, CategorisedRow, Method
from categoriser.sheets import ensure_categorised_sheet, ensure_issues_sheet, record_issues, upsert_categorised
from fakes import FakeSpreadsheet, FakeWorksheet

SETTINGS = Settings(spreadsheet_id="x", credentials_path=Path("/dev/null"))
HEADER = list(CATEGORISED_COLUMNS)
TXN_HEADER = ["transactionId", "amount", "category", "classification"]


def row(tid, category="Groceries"):
    return CategorisedRow(tid, "joint", "Essentials", category, "", Method.RULE, 0.9, "R1", False,
                          "2026-10-31T06:15:00Z")


def sheet_row(tid, reviewed=False, category="Old"):
    return [tid, "joint", "Essentials", category, "", "rule", 1, "R1", reviewed, "2026-10-01T00:00:00Z"]


def test_creates_missing_sheet_with_header_freeze_and_checkboxes():
    ss = FakeSpreadsheet({"transactions": FakeWorksheet("transactions", [TXN_HEADER])})
    ws = ensure_categorised_sheet(ss, SETTINGS)
    assert ws.title == "categorised"
    assert ws.values == [HEADER]
    assert ws.frozen_rows == 1
    rule = ss.requests[0]["setDataValidation"]
    assert rule["range"]["startColumnIndex"] == REVIEWED_COLUMN
    assert rule["rule"]["condition"]["type"] == "BOOLEAN"


def test_fills_header_of_an_empty_existing_sheet():
    ss = FakeSpreadsheet({"categorised": FakeWorksheet("categorised")})
    assert ensure_categorised_sheet(ss, SETTINGS).values == [HEADER]


def test_existing_sheet_with_other_headers_is_left_alone():
    ws = FakeWorksheet("categorised", [["id", "notes"]])
    with pytest.raises(ValueError):
        ensure_categorised_sheet(FakeSpreadsheet({"categorised": ws}), SETTINGS)
    assert ws.values == [["id", "notes"]]


def test_upsert_updates_appends_and_keeps_reviewed_rows():
    ws = FakeWorksheet("categorised", [HEADER, sheet_row("a"), sheet_row("b", reviewed=True)])
    plan = upsert_categorised(ws, SETTINGS, [row("a"), row("b"), row("c")])
    assert ws.values[1][3] == "Groceries"          # a updated
    assert ws.values[2][3] == "Old"                # b reviewed: untouched
    assert ws.values[3][0] == "c"                  # c appended
    assert plan.skipped_reviewed == ["b"]


def test_row_ticked_between_plan_and_write_is_not_overwritten():
    ws = FakeWorksheet("categorised", [HEADER, sheet_row("a")])
    ws.before_batch_get = lambda w: w.values[1].__setitem__(REVIEWED_COLUMN, True)
    plan = upsert_categorised(ws, SETTINGS, [row("a")])
    assert ws.values[1][3] == "Old" and ws.values[1][REVIEWED_COLUMN] is True
    assert plan.updates == {} and plan.skipped_reviewed == ["a"]
    assert "batch_update" not in ws.calls


def test_row_moved_between_plan_and_write_is_skipped():
    ws = FakeWorksheet("categorised", [HEADER, sheet_row("a"), sheet_row("b")])
    ws.before_batch_get = lambda w: w.values.__setitem__(slice(1, 3), [w.values[2], w.values[1]])
    plan = upsert_categorised(ws, SETTINGS, [row("a")])
    assert plan.skipped_moved == ["a"]
    assert [r[3] for r in ws.values[1:]] == ["Old", "Old"]


def test_never_writes_to_transactions():
    ws = FakeWorksheet("transactions", [TXN_HEADER])
    with pytest.raises(PermissionError):
        upsert_categorised(ws, SETTINGS, [row("a")])
    assert ws.calls == []


def test_duplicate_ids_are_skipped_and_the_rest_is_written():
    ws = FakeWorksheet("categorised", [HEADER, sheet_row("a"), sheet_row("b"), sheet_row("a")])
    plan = upsert_categorised(ws, SETTINGS, [row("a"), row("b")])
    assert plan.duplicates == {"a": [2, 4]}
    assert [r[3] for r in ws.values[1:]] == ["Old", "Groceries", "Old"]


def test_issues_log_records_once_then_resolves():
    ss = FakeSpreadsheet()
    ws = ensure_issues_sheet(ss, SETTINGS)
    assert ws.values == [list(ISSUE_COLUMNS)] and ss.requests == []  # no checkbox column

    found = duplicate_issues({"a": [2, 4]}, "categorised")
    record_issues(ws, SETTINGS, found, {DUPLICATE_ID}, "2026-10-31T06:15:00Z")
    record_issues(ws, SETTINGS, found, {DUPLICATE_ID}, "2026-11-01T06:15:00Z")
    assert len(ws.values) == 2                                       # logged once
    assert ws.values[1][COL["lastSeen"]] == "2026-11-01T06:15:00Z"

    plan = record_issues(ws, SETTINGS, [], {DUPLICATE_ID}, "2026-11-02T06:15:00Z")
    assert plan.resolved == 1 and ws.values[1][COL["status"]] == "resolved"


def test_issues_log_is_never_written_to_transactions():
    ws = FakeWorksheet("transactions", [TXN_HEADER])
    with pytest.raises(PermissionError):
        record_issues(ws, SETTINGS, duplicate_issues({"a": [2, 3]}, "categorised"), {DUPLICATE_ID}, "t")
    assert ws.calls == []
