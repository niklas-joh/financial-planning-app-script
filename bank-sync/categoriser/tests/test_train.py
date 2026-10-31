from datetime import date
from pathlib import Path

import pytest

from categoriser import train
from categoriser.config import Settings
from categoriser.evaluate import evaluate, write_report
from categoriser.legacy import LegacyRow
from categoriser.model import load
from builders import TAXONOMY, txn

GROCERIES = ("Essentials", "Groceries", "Supermarket")
CAFE = ("Wants/Pleasure", "Eating out", "Cafe")


def inputs(labels):
    """Bank Sync rows plus legacy rows joined to them (same amount and day)."""
    transactions, legacy = [], []
    for i, (name, label) in enumerate(labels):
        day = date(2026, 1 + i % 9, 1 + i % 27)
        transactions.append(txn(counterparty=name, amount=-(i + 1), day=day.isoformat()))
        legacy.append(LegacyRow(day, -(i + 1), "Old", *label, counterparty=name, description=""))
    return transactions, TAXONOMY, [], legacy, {}


@pytest.fixture
def home(tmp_path, monkeypatch):
    monkeypatch.setenv("CATEGORISER_HOME", str(tmp_path))
    monkeypatch.setattr(train, "load_settings", lambda: Settings("x", Path("/dev/null")))
    monkeypatch.setattr(train, "models_dir", lambda: tmp_path / "models")
    monkeypatch.setattr("sys.argv", ["train"])
    return tmp_path


def test_train_saves_a_loadable_model_with_holdout_scores(home, monkeypatch):
    labels = [("Acme Market", GROCERIES)] * 6 + [("Corner Cafe", CAFE)] * 6
    monkeypatch.setattr(train, "load_inputs", lambda _settings: inputs(labels))
    train.main()

    model, status = load(home / "models")
    assert status == "loaded"
    assert model.meta["examplesByKind"] == {"joined": 12}
    assert model.meta["holdout"]["minConfidence"] == 0.9 and "0.9" in model.meta["holdout"]["model"]
    assert len(list((home / "data" / "eval").glob("holdout-*.md"))) == 1


def test_train_with_too_few_labels_exits_with_a_message(home, monkeypatch):
    monkeypatch.setattr(train, "load_inputs", lambda _settings: inputs([("Acme Market", GROCERIES)] * 3))
    with pytest.raises(SystemExit, match="No model saved"):
        train.main()
    assert not (home / "models").exists()


def test_write_report_goes_to_the_runtime_data_folder(home):
    transactions, taxonomy, manual, legacy, reviewed = inputs([("Acme Market", GROCERIES)] * 3)
    ev = evaluate(transactions, legacy, reviewed, taxonomy, manual, "common", date(2026, 3, 1))
    path = write_report(ev, transactions, legacy, reviewed, taxonomy, manual, "common")
    assert path.parent == home / "data" / "eval" and "# Holdout evaluation" in path.read_text()


def test_load_inputs_reads_the_job_sheets_and_the_legacy_sheet_read_only(monkeypatch):
    from categoriser import sheets
    from categoriser.evaluate import load_inputs
    from categoriser.schema import CATEGORISED_COLUMNS
    from categoriser.taxonomy import TAXONOMY_COLUMNS
    from builders import rule_row, rules_sheet
    from fakes import FakeSpreadsheet, FakeWorksheet

    tx = FakeWorksheet("transactions", [["transactionId"]])
    tx.get_all_records = lambda **_: [txn(counterparty="Acme Market", tid="a", owner="alice")]
    ss = FakeSpreadsheet({
        "transactions": tx,
        "categorised": FakeWorksheet("categorised", [list(CATEGORISED_COLUMNS),
                                                     ["a", "alice", *GROCERIES, "rule", "", "R1", True, ""]]),
        "categoriser_taxonomy": FakeWorksheet("categoriser_taxonomy", [list(TAXONOMY_COLUMNS), list(GROCERIES)]),
        "categoriser_rules": FakeWorksheet("categoriser_rules", rules_sheet(rule_row("R1", counterpartyContains="acme"),
                                                                         rule_row("m-1", source="mined", counterpartyContains="acme"))),
    })
    legacy_tab = FakeWorksheet("All Transactions", [
        ["Source", "Type", "Category", "Sub-Category", "Date", "Amount", "Beneficiary Name", "Operation"],
        ["Old", *GROCERIES, "2025-01-02", -5, "Acme Market", "Card"]])
    opened = []
    monkeypatch.setattr(sheets, "open_spreadsheet", lambda _s: ss)
    monkeypatch.setattr(sheets, "open_readonly", lambda _s, sid: opened.append(sid) or FakeSpreadsheet({"All Transactions": legacy_tab}))

    transactions, taxonomy, manual, legacy, reviewed = load_inputs(
        Settings("x", Path("/dev/null"), legacy_spreadsheet_id="legacy-id"))
    assert [t["transactionId"] for t in transactions] == ["a"] and len(taxonomy) == 1
    assert [r.ruleId for r in manual] == ["R1"] and len(legacy) == 1
    assert reviewed == {"a": GROCERIES} and opened == ["legacy-id"]
    with pytest.raises(SystemExit):
        load_inputs(Settings("x", Path("/dev/null")))
