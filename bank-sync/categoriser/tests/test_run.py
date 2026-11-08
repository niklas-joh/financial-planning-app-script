import json
from pathlib import Path

import pytest

from categoriser import run
from categoriser.config import Settings
from categoriser.issues import COL
from categoriser.model import MODEL_FILE, save, train
from categoriser.rules import RCOL, RULE_COLUMNS
from categoriser.review import REVIEW_COLUMNS, VCOL
from categoriser.schema import CATEGORISED_COLUMNS
from categoriser.taxonomy import TAXONOMY_COLUMNS
from builders import TAXONOMY, review_row, review_sheet, rule_row, txn
from fakes import FakeSpreadsheet, FakeWorksheet

HEADER = list(CATEGORISED_COLUMNS)
TAXONOMY_ROWS = [list(TAXONOMY_COLUMNS), ["Essentials", "Groceries", "Supermarket"],
                 ["Internal Transfer", "Internal Transfer", "From Alice"]]


def sheet_row(tid, method="rule"):
    return [tid, "joint", "Essentials", "Old", "", method, 1, "R1", False, "2026-10-01T00:00:00Z"]


def spreadsheet(transactions, categorised, rules=None, taxonomy=None, review=None):
    tx = FakeWorksheet("transactions", [["transactionId"]])
    tx.get_all_records = lambda **_: transactions
    sheets = {"transactions": tx, "categorised": FakeWorksheet("categorised", categorised)}
    if rules is not None:
        sheets["categoriser_rules"] = FakeWorksheet("categoriser_rules", rules)
    if taxonomy is not None:
        sheets["categoriser_taxonomy"] = FakeWorksheet("categoriser_taxonomy", taxonomy)
    if review is not None:
        sheets["categoriser_review"] = FakeWorksheet("categoriser_review", review)
    return FakeSpreadsheet(sheets)


@pytest.fixture(autouse=True)
def models(tmp_path, monkeypatch):
    """Never the real ~/categoriser/models."""
    monkeypatch.setattr(run, "models_dir", lambda: tmp_path)
    return tmp_path


@pytest.fixture(autouse=True)
def data(tmp_path, monkeypatch):
    """Never the real ~/categoriser/data."""
    folder = tmp_path / "data"
    monkeypatch.setattr(run, "data_dir", lambda: folder)
    return folder


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
    assert rules_ws.values[1][RCOL["hits"]:] == [1, "2026-05-02", 1, "Acme Market (1)"]
    assert rows["g"][10:] == ["2026-05-02", "Acme Market", "", -12]   # context copied from transactions
    issues = ss.sheets["categoriser_issues"].values
    # "stale" is not on transactions, so it is also reported as orphaned
    assert [r[COL["issueKey"]] for r in issues[1:]] == ["invalid_rule:R2", "orphaned_categorised_row:stale"]
    assert "by_rule=1 by_transfer=2 uncovered=1" in caplog.text and "cleared=1" in caplog.text


def test_renamed_taxonomy_header_stops_the_run_before_any_write(monkeypatch):
    before = [HEADER, sheet_row("a")]
    ss = spreadsheet([txn(tid="a")], [list(r) for r in before], [list(RULE_COLUMNS)], [["kind", "category", "x"]])
    with pytest.raises(ValueError):
        run_with(monkeypatch, ss)
    assert ss.sheets["categorised"].values == before


def model_spreadsheet(categorised):
    acme = txn("bob", -25, counterparty="Acme Market", tid="new")
    card = txn("bob", -25, counterparty="Acme Market", tid="card", classification="card")
    taxonomy = TAXONOMY_ROWS + [["Wants/Pleasure", "Eating out", "Cafe"]]
    return spreadsheet([acme, card], categorised, [list(RULE_COLUMNS)], taxonomy)


def trained(models):
    grocery, cafe = ("Essentials", "Groceries", "Supermarket"), ("Wants/Pleasure", "Eating out", "Cafe")
    examples = [(txn(counterparty="Acme Market", amount=-20 - i), grocery, 1.0) for i in range(5)]
    examples += [(txn(counterparty="Corner Cafe", amount=-3 - i), cafe, 1.0) for i in range(5)]
    save(train(examples, TAXONOMY, "2026-10-30T00:00:00Z"), models)


def test_model_labels_confident_uncovered_rows_and_clears_its_stale_rows(monkeypatch, models, caplog):
    caplog.set_level("INFO", logger="categoriser")
    trained(models)
    monkeypatch.setattr(run, "load_settings", lambda: Settings("x", Path("/dev/null"), model_min_confidence=0.5))
    monkeypatch.setattr(run, "open_spreadsheet", lambda _: ss)
    ss = model_spreadsheet([HEADER, sheet_row("old-model", method="model")])
    run.main()

    rows = {r[0]: r for r in ss.sheets["categorised"].values[1:]}
    assert rows["new"][1:6] == ["bob", "Essentials", "Groceries", "Supermarket", "model"]
    assert 0.5 <= rows["new"][6] <= 1 and rows["new"][7] == ""
    assert "card" not in rows                                   # card rows are left to rules
    assert rows["old-model"][1:8] == [""] * 7                   # the model no longer labels it
    assert "model=loaded by_model=1" in caplog.text


def test_broken_model_skips_the_model_step_and_keeps_its_rows(monkeypatch, models, caplog):
    caplog.set_level("INFO", logger="categoriser")
    trained(models)
    (models / MODEL_FILE).write_bytes(b"broken")
    before = sheet_row("old-model", method="model")
    ss = model_spreadsheet([HEADER, list(before)])
    run_with(monkeypatch, ss)

    assert ss.sheets["categorised"].values == [HEADER, before]
    assert "model=unreadable by_model=0" in caplog.text


# --- review inbox -------------------------------------------------------------

GROCERY = "Essentials › Groceries › Supermarket"


def review_spreadsheet(answer=GROCERY, extra_rows=()):
    txns = [txn("bob", -5, counterparty="Acme Market", tid="a1", day="2026-05-01"),
            txn("bob", -6, counterparty="Acme Market", tid="a2", day="2026-05-02"),
            txn("bob", -7, counterparty="Lone Shop", tid="n")]
    review = review_sheet(review_row("a1", answer=answer, accept=True, apply_to="Same merchant", merchant_text="acme market"),
                          *extra_rows)
    return spreadsheet(txns, [HEADER], [list(RULE_COLUMNS)], TAXONOMY_ROWS, review)


def inbox_ids(ss):
    return [r[0] for r in ss.sheets["categoriser_review"].values[1:] if r and r[0]]


def history(data):
    path = data / "review-history.jsonl"
    return [json.loads(line) for line in path.read_text().splitlines()] if path.exists() else []


def test_answer_for_the_same_merchant_labels_the_merchant_in_the_same_run(monkeypatch, data, caplog):
    caplog.set_level("INFO", logger="categoriser")
    ss = review_spreadsheet()
    run_with(monkeypatch, ss)

    rows = {r[0]: r for r in ss.sheets["categorised"].values[1:]}
    [rule] = ss.sheets["categoriser_rules"].values[1:]
    rid = rule[RCOL["ruleId"]]
    assert rid.startswith("r-") and rule[RCOL["source"]] == "review" and rule[RCOL["active"]] is True
    assert rows["a1"][1:6] == ["bob", "Essentials", "Groceries", "Supermarket", "manual"]
    assert rows["a1"][CATEGORISED_COLUMNS.index("reviewed")] is True and rows["a1"][7] == rid
    assert rows["a2"][5] == "rule" and rows["a2"][7] == rid
    assert inbox_ids(ss) == ["n"]
    assert [h["transactionId"] for h in history(data)] == ["a1"]
    assert "review_answers=1" in caplog.text and "review_rules_added=1" in caplog.text
    assert "review=written" in caplog.text


def test_second_run_changes_nothing(monkeypatch):
    ss = review_spreadsheet()
    run_with(monkeypatch, ss)
    names = ("categorised", "categoriser_rules", "categoriser_review")
    after_first = {n: [list(r) for r in ss.sheets[n].values] for n in names}
    run_with(monkeypatch, ss)
    assert {n: ss.sheets[n].values for n in names} == after_first


def test_answer_seen_twice_is_applied_once(monkeypatch, data, caplog):
    caplog.set_level("INFO", logger="categoriser")
    ss = review_spreadsheet(extra_rows=[review_row("n")])
    review_ws = ss.sheets["categoriser_review"]

    def tick_n(ws):
        ws.values[2][VCOL["accept"]] = True
    review_ws.before_batch_get = tick_n
    run_with(monkeypatch, ss)
    assert "review=skipped_busy" in caplog.text and inbox_ids(ss) == ["a1", "n"]

    review_ws.before_batch_get = None
    run_with(monkeypatch, ss)
    assert len(ss.sheets["categoriser_rules"].values) == 2 and len(history(data)) == 1
    a1 = next(r for r in ss.sheets["categorised"].values[1:] if r[0] == "a1")
    assert a1[5] == "manual" and a1[CATEGORISED_COLUMNS.index("reviewed")] is True


def test_invalid_answer_is_logged_and_stays_ticked(monkeypatch):
    ss = review_spreadsheet(answer="Essentials › Groceries › Bakery")
    run_with(monkeypatch, ss)
    issues = {r[COL["issueKey"]]: r for r in ss.sheets["categoriser_issues"].values[1:]}
    assert issues["invalid_review_answer:a1"][COL["status"]] == "open"
    review = {r[0]: r for r in ss.sheets["categoriser_review"].values[1:] if r and r[0]}
    assert review["a1"][VCOL["accept"]] is True
    assert "a1" not in {r[0] for r in ss.sheets["categorised"].values[1:]}
    assert len(ss.sheets["categoriser_rules"].values) == 1


def test_inbox_is_created_and_filled_on_the_first_run(monkeypatch):
    ss = spreadsheet([txn(counterparty="Acme Market", tid="a"), txn(counterparty="Lone Shop", tid="l")],
                     [HEADER], [list(RULE_COLUMNS)], TAXONOMY_ROWS)
    run_with(monkeypatch, ss)
    review = ss.sheets["categoriser_review"].values
    assert review[0] == list(REVIEW_COLUMNS) and sorted(inbox_ids(ss)) == ["a", "l"]


def test_answered_spot_check_is_not_offered_again(monkeypatch, models):
    trained(models)
    monkeypatch.setattr(run, "load_settings", lambda: Settings("x", Path("/dev/null"), model_min_confidence=0.5))
    review = review_sheet(review_row("s", answer=GROCERY, suggested=GROCERY, accept=True, spot_check=True))
    ss = spreadsheet([txn("bob", -25, counterparty="Acme Market", tid="s")], [HEADER], [list(RULE_COLUMNS)],
                     TAXONOMY_ROWS, review)
    monkeypatch.setattr(run, "open_spreadsheet", lambda _: ss)
    run.main()
    s = next(r for r in ss.sheets["categorised"].values[1:] if r[0] == "s")
    assert s[5] == "manual" and inbox_ids(ss) == []


def test_failing_guess_for_the_inbox_does_not_stop_the_run(monkeypatch, caplog):
    class Broken:
        def rows(self, *_):
            return []

        def predict(self, _):
            raise RuntimeError("boom")

    monkeypatch.setattr(run, "load_model", lambda _: (Broken(), "loaded"))
    ss = spreadsheet([txn(counterparty="Acme Market", tid="a")], [HEADER], [list(RULE_COLUMNS)], TAXONOMY_ROWS)
    run_with(monkeypatch, ss)
    [row] = [r for r in ss.sheets["categoriser_review"].values[1:] if r and r[0]]
    assert row[0] == "a" and row[VCOL["answer"]] == ""
    assert len(ss.sheets["categoriser_issues"].values) == 1   # the issues step still ran (header only)


def test_categorised_grows_past_its_first_1000_rows_and_counts_coverage(monkeypatch, caplog):
    """A new tab has 1000 rows whose unticked checkboxes read as FALSE; every labelled row must still land."""
    caplog.set_level("INFO", logger="categoriser")
    blank = [""] * len(CATEGORISED_COLUMNS)
    blank[CATEGORISED_COLUMNS.index("reviewed")] = False
    shops = [txn("alice", -5, counterparty="Acme Market", tid=f"s{i}") for i in range(1500)]
    rules = [list(RULE_COLUMNS), rule_row("R1", counterpartyContains="acme market")]
    ss = spreadsheet(shops + [txn("bob", -1, counterparty="Nobody", tid="n")],
                     [HEADER, sheet_row("gone")] + [list(blank) for _ in range(998)], rules, TAXONOMY_ROWS)
    run_with(monkeypatch, ss)

    written = ss.sheets["categorised"].values
    ids = {r[0] for r in written[1:] if r and r[0]}
    assert {f"s{i}" for i in range(1500)} <= ids
    assert ss.sheets["categorised"].row_count >= len(written)
    issues = ss.sheets["categoriser_issues"].values[1:]
    assert [(r[COL["type"]], r[COL["transactionId"]]) for r in issues] == [("orphaned_categorised_row", "gone")]
    assert "transactions=1501 categorised=1500 uncategorised=1 orphaned=1 not_written=0" in caplog.text


def test_merchant_and_description_rule_labels_only_rows_with_both(monkeypatch):
    txns = [txn("bob", -5, counterparty="Trading Markets", description="Savings for Alice", tid="a1", day="2026-05-01"),
            txn("bob", -6, counterparty="Trading Markets", description="Savings for Alice", tid="a2", day="2026-05-02"),
            txn("bob", -7, counterparty="Trading Markets", description="Savings for Carol", tid="c1", day="2026-05-03")]
    review = review_sheet(review_row("a1", answer=GROCERY, accept=True, apply_to="Merchant + description",
                                     merchant_text="trading markets", description_text="savings for alice"))
    ss = spreadsheet(txns, [HEADER], [list(RULE_COLUMNS)], TAXONOMY_ROWS, review)
    run_with(monkeypatch, ss)
    [rule] = ss.sheets["categoriser_rules"].values[1:]
    assert (rule[RCOL["counterpartyContains"]], rule[RCOL["descriptionContains"]]) == \
        ("trading markets", "savings for alice")
    labelled = {r[0]: r[5] for r in ss.sheets["categorised"].values[1:] if r[2]}
    assert labelled == {"a1": "manual", "a2": "rule"}
    assert inbox_ids(ss) == ["c1"]
