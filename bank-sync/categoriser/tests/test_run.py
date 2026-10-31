from pathlib import Path

from categoriser import run
from categoriser.config import Settings
from categoriser.issues import COL
from categoriser.schema import CATEGORISED_COLUMNS
from fakes import FakeSpreadsheet, FakeWorksheet

HEADER = list(CATEGORISED_COLUMNS)


def sheet_row(tid):
    return [tid, "joint", "Essentials", "Old", "", "rule", 1, "R1", False, "2026-10-01T00:00:00Z"]


def test_daily_run_continues_past_duplicates_and_logs_them_once(monkeypatch, caplog):
    caplog.set_level("INFO", logger="categoriser")
    transactions = FakeWorksheet("transactions", [["transactionId", "amount"], ["a", -5], ["b", -7]])
    categorised = FakeWorksheet("categorised", [HEADER, sheet_row("a"), sheet_row("a")])
    ss = FakeSpreadsheet({"transactions": transactions, "categorised": categorised})
    transactions.get_all_records = lambda **_: [{"transactionId": "a"}, {"transactionId": "b"}]
    monkeypatch.setattr(run, "load_settings", lambda: Settings("x", Path("/dev/null")))
    monkeypatch.setattr(run, "open_spreadsheet", lambda _: ss)

    run.main()
    run.main()

    issues = ss.sheets["categoriser_issues"].values
    assert len(issues) == 2 and issues[1][COL["transactionId"]] == "a"
    assert transactions.calls == []                      # transactions never written
    assert "skipped_duplicate=1" in caplog.text
