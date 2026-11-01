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


def test_sheet_from_before_the_context_columns_gets_them_added_at_the_end():
    old = HEADER[: HEADER.index("updatedAt") + 1]
    row = sheet_row("a")
    ws = FakeWorksheet("categorised", [old, row], col_count=len(old))
    ensure_categorised_sheet(FakeSpreadsheet({"categorised": ws}), SETTINGS)
    assert ws.values == [HEADER, row] and ws.col_count == len(HEADER)


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


# --- rules and taxonomy sheets ---------------------------------------------

from categoriser.engine import RuleStats  # noqa: E402
from categoriser.rules import RCOL, RULE_COLUMNS  # noqa: E402
from categoriser.sheets import append_rows, ensure_rules_sheet, ensure_taxonomy_sheet, write_rule_stats  # noqa: E402
from categoriser.taxonomy import TAXONOMY_COLUMNS  # noqa: E402


def test_creates_rules_sheet_with_checkbox_and_dropdowns():
    ss = FakeSpreadsheet()
    ws = ensure_rules_sheet(ss, SETTINGS)
    assert ws.values == [list(RULE_COLUMNS)] and ws.frozen_rows == 1
    rules = {r["setDataValidation"]["range"]["startColumnIndex"]: r["setDataValidation"]["rule"]
             for r in ss.requests if "setDataValidation" in r}
    assert rules[RCOL["active"]]["condition"]["type"] == "BOOLEAN"
    assert [v["userEnteredValue"] for v in rules[RCOL["direction"]]["condition"]["values"]] == ["in", "out", "any"]
    assert rules[RCOL["type"]]["condition"]["values"][0]["userEnteredValue"] == "='categoriser_taxonomy'!A2:A"
    assert rules[RCOL["subCategory"]]["condition"]["values"][0]["userEnteredValue"] == "='categoriser_taxonomy'!C2:C"
    assert rules[RCOL["owner"]]["condition"]["values"][0]["userEnteredValue"] == "='account_owners'!A2:A"
    assert rules[RCOL["owner"]]["strict"] is False
    assert {"setBasicFilter": {"filter": {"range": {"sheetId": ws.id}}}} in ss.requests  # sortable


def test_rules_sheet_filter_is_added_once_and_an_existing_one_is_kept():
    ss = FakeSpreadsheet({"categoriser_rules": FakeWorksheet("categoriser_rules", rules_values("R1"))})
    ensure_rules_sheet(ss, SETTINGS)
    ensure_rules_sheet(ss, SETTINGS)
    assert sum("setBasicFilter" in r for r in ss.requests) == 1


def test_rules_sheet_columns_can_be_reordered_and_missing_job_columns_are_added():
    header = ["type", "category", "subCategory", "ruleId", "active", "priority", "account", "direction",
              "counterpartyContains", "descriptionContains", "textRegex", "providerCategory", "amountMin",
              "amountMax", "owner", "notes", "source", "my own column", "hits", "lastMatched"]
    ss = FakeSpreadsheet({"categoriser_rules": FakeWorksheet("categoriser_rules", [header], col_count=len(header))})
    ws = ensure_rules_sheet(ss, SETTINGS)
    assert ws.values[0] == [*header, "wins", "sample"] and ws.col_count == len(header) + 2


def test_rules_sheet_with_a_repeated_column_stops_the_run():
    header = [*RULE_COLUMNS, "hits"]
    ss = FakeSpreadsheet({"categoriser_rules": FakeWorksheet("categoriser_rules", [header])})
    with pytest.raises(ValueError, match="unexpected headers"):
        ensure_rules_sheet(ss, SETTINGS)


def test_reordered_rules_sheet_gets_validations_and_stats_by_column_name():
    header = list(reversed(RULE_COLUMNS))
    col = {n: i for i, n in enumerate(header)}
    row = [""] * len(header)
    row[col["ruleId"]], row[col["notes"]] = "R1", "keep me"
    ss = FakeSpreadsheet({"categoriser_rules": FakeWorksheet("categoriser_rules", [header, row], row_count=2)})
    ws = ss.sheets["categoriser_rules"]
    write_rule_stats(ws, SETTINGS, {"R1": RuleStats(5, 2, "2026-10-01", "Acme Market (5)")})
    got = {n: ws.values[1][col[n]] for n in ("ruleId", "notes", "hits", "wins", "lastMatched", "sample")}
    assert got == {"ruleId": "R1", "notes": "keep me", "hits": 5, "wins": 2, "lastMatched": "2026-10-01",
                   "sample": "Acme Market (5)"}
    new = [""] * len(header)
    new[col["ruleId"]] = "R2"
    append_rows(ws, SETTINGS, [new], key_column=col["ruleId"] + 1)   # grows the grid, re-applies dropdowns
    checkbox = [r["setDataValidation"]["range"]["startColumnIndex"] for r in ss.requests
                if r["setDataValidation"]["rule"]["condition"]["type"] == "BOOLEAN"]
    assert checkbox == [col["active"]] and ws.values[2][col["ruleId"]] == "R2"


def test_creates_taxonomy_sheet_and_appends_rows():
    ss = FakeSpreadsheet()
    ws = ensure_taxonomy_sheet(ss, SETTINGS)
    append_rows(ws, SETTINGS, [["A", "B", "C"]])
    assert ws.values == [list(TAXONOMY_COLUMNS), ["A", "B", "C"]]


def test_append_rows_refuses_other_sheets():
    with pytest.raises(PermissionError):
        append_rows(FakeWorksheet("transactions"), SETTINGS, [["x"]])


def rules_values(*ids):
    rows = [list(RULE_COLUMNS)]
    for rid in ids:
        row = [""] * len(RULE_COLUMNS)
        row[0], row[RCOL["notes"]] = rid, f"note {rid}"
        rows.append(row)
    return rows


def test_rule_stats_find_rows_by_id_and_touch_only_the_job_columns():
    ws = FakeWorksheet("categoriser_rules", rules_values("R1", "R2", "dup", "dup"))
    ws.values[2][RCOL["hits"]:] = [4, "2026-01-01", 4, "Acme (4)"]
    changed = write_rule_stats(ws, SETTINGS, {
        "R1": RuleStats(3, 2, "2026-10-01", "Acme (3)"), "R2": RuleStats(4, 4, "2026-01-01", "Acme (4)"),
        "dup": RuleStats(1, 1, "x", ""), "gone": RuleStats(9, 9, "y", ""),
    })
    assert changed == 1                                     # R2 unchanged, dup ambiguous, gone missing
    assert ws.values[1][RCOL["hits"]:] == [3, "2026-10-01", 2, "Acme (3)"]
    assert ws.values[1][RCOL["notes"]] == "note R1"
    assert ws.values[3][RCOL["hits"]] == "" and ws.calls.count("batch_update") == 1


def test_rule_stats_with_nothing_to_change_does_not_write():
    ws = FakeWorksheet("categoriser_rules", rules_values("R1"))
    assert write_rule_stats(ws, SETTINGS, {}) == 0 and "batch_update" not in ws.calls


def test_appends_go_below_the_last_key_not_below_blank_checkbox_rows():
    ss = FakeSpreadsheet({"categoriser_rules": FakeWorksheet("categoriser_rules", rules_values("R1"), row_count=10)})
    ws = ss.sheets["categoriser_rules"]
    ws.values += [["", False] for _ in range(8)]              # unchecked boxes read as FALSE
    append_rows(ws, SETTINGS, [["R2"], ["R3"]])
    assert [r[0] for r in ws.values[:4]] == ["ruleId", "R1", "R2", "R3"]
    assert "add_rows" not in ws.calls


def test_appends_past_the_grid_add_rows_and_reapply_validations():
    ss = FakeSpreadsheet({"categoriser_rules": FakeWorksheet("categoriser_rules", rules_values("R1"), row_count=2)})
    ws = ss.sheets["categoriser_rules"]
    append_rows(ws, SETTINGS, [["R2"], ["R3"]])
    assert ws.row_count == 4 and [r[0] for r in ws.values] == ["ruleId", "R1", "R2", "R3"]
    assert {r["setDataValidation"]["range"]["startColumnIndex"] for r in ss.requests} >= {RCOL["active"]}


def test_backfill_context_writes_only_the_context_cells_once():
    from categoriser.sheets import backfill_context
    from builders import txn
    reviewed = ["a", "joint", "Essentials", "Mine", "", "manual", "", "", True, "2026-01-01T00:00:00Z"]
    ws = FakeWorksheet("categorised", [HEADER, list(reviewed), ["b", "joint"]])
    txns = [txn(tid="a", counterparty="Acme Market", day="2026-05-02", amount=-3), txn(tid="b")]
    assert backfill_context(ws, SETTINGS, txns) == 2
    assert ws.values[1] == [*reviewed, "2026-05-02", "Acme Market", "", -3]   # labels untouched
    assert backfill_context(ws, SETTINGS, txns) == 0                           # nothing blank any more


def test_backfill_context_refuses_other_sheets():
    from categoriser.sheets import backfill_context
    with pytest.raises(PermissionError):
        backfill_context(FakeWorksheet("transactions"), SETTINGS, [])
