from pathlib import Path

import pytest

from categoriser import run
from categoriser.config import Settings
from categoriser.issues import COL
from categoriser.rules import RCOL, RULE_COLUMNS
from categoriser.schema import CATEGORISED_COLUMNS
from categoriser.taxonomy import TAXONOMY_COLUMNS
from builders import rule_row, txn
from fakes import FakeSpreadsheet, FakeWorksheet

HEADER = list(CATEGORISED_COLUMNS)
TAXONOMY_ROWS = [list(TAXONOMY_COLUMNS), ["Essentials", "Groceries", "Supermarket"],
                 ["Internal Transfer", "Internal Transfer", "From Alice"]]


def sheet_row(tid, method="rule"):
    return [tid, "joint", "Essentials", "Old", "", method, 1, "R1", False, "2026-10-01T00:00:00Z"]


def spreadsheet(transactions, categorised, rules=None, taxonomy=None):
    tx = FakeWorksheet("transactions", [["transactionId"]])
    tx.get_all_records = lambda **_: transactions
    sheets = {"transactions": tx, "categorised": FakeWorksheet("categorised", categorised)}
    if rules is not None:
        sheets["categoriser_rules"] = FakeWorksheet("categoriser_rules", rules)
    if taxonomy is not None:
        sheets["categoriser_taxonomy"] = FakeWorksheet("categoriser_taxonomy", taxonomy)
    return FakeSpreadsheet(sheets)


def run_with(monkeypatch, ss):
    monkeypatch.setattr(run, "load_settings", lambda: Settings("x", Path("/dev/null")))
    monkeypatch.setattr(run, "open_spreadsheet", lambda _: ss)
    run.main()


def test_daily_run_continues_past_duplicates_and_logs_them_once(monkeypatch, caplog):
    caplog.set_level("INFO", logger="categoriser")
    ss = spreadsheet([{"transactionId": "a"}, {"transactionId": "b"}], [HEADER, sheet_row("a"), sheet_row("a")])
    run_with(monkeypatch, ss)
    run_with(monkeypatch, ss)

    issues = ss.sheets["categoriser_issues"].values
    assert len(issues) == 2 and issues[1][COL["transactionId"]] == "a"
    assert ss.sheets["transactions"].calls == []         # transactions never written
    assert "skipped_duplicate=1" in caplog.text
    assert ss.sheets["categoriser_rules"].values == [list(RULE_COLUMNS)]  # created on first run


def test_rules_and_transfers_reach_categorised_and_rule_stats(monkeypatch, caplog):
    caplog.set_level("INFO", logger="categoriser")
    grocery = txn("alice", -12, counterparty="Acme Market", tid="g", day="2026-05-02")
    out, inc = txn("alice", -300, tid="o"), txn("common", 300, tid="i")
    rules = [list(RULE_COLUMNS), rule_row("R1", counterpartyContains="acme market"),
             rule_row("R2", active=True, textRegex="(")]
    ss = spreadsheet([grocery, out, inc, txn("bob", -1, counterparty="Nobody", tid="n")],
                     [HEADER, sheet_row("stale")], rules, TAXONOMY_ROWS)
    run_with(monkeypatch, ss)

    rows = {r[0]: r for r in ss.sheets["categorised"].values[1:]}
    assert rows["g"][1:5] == ["alice", "Essentials", "Groceries", "Supermarket"] and rows["g"][7] == "R1"
    assert rows["o"][4] == rows["i"][4] == "From Alice" and rows["o"][7] == "transfer:pair"
    assert "n" not in rows
    assert rows["stale"][1:8] == [""] * 7                  # its rule no longer covers it
    rules_ws = ss.sheets["categoriser_rules"]
    assert rules_ws.values[1][RCOL["hits"]:] == [1, "2026-05-02"]
    issues = ss.sheets["categoriser_issues"].values
    assert [r[COL["issueKey"]] for r in issues[1:]] == ["invalid_rule:R2"]
    assert "by_rule=1 by_transfer=2 uncovered=1" in caplog.text and "cleared=1" in caplog.text


def test_renamed_taxonomy_header_stops_the_run_before_any_write(monkeypatch):
    before = [HEADER, sheet_row("a")]
    ss = spreadsheet([txn(tid="a")], [list(r) for r in before], [list(RULE_COLUMNS)], [["kind", "category", "x"]])
    with pytest.raises(ValueError):
        run_with(monkeypatch, ss)
    assert ss.sheets["categorised"].values == before
