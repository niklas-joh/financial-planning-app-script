from datetime import date

from categoriser.evaluate import coverage_by_account, evaluate, render
from categoriser.legacy import LegacyRow
from builders import TAXONOMY, txn

GROCERIES = ("Essentials", "Groceries", "Supermarket")
CAFE = ("Wants/Pleasure", "Eating out", "Cafe")


def labelled(day, amount, label, name):
    return LegacyRow(day, amount, "Old", *label, counterparty=name, description="")


def test_rules_are_mined_before_the_cutoff_and_scored_after():
    train = [txn(counterparty="Acme", amount=-i, day=f"2026-01-0{i}") for i in range(1, 4)]
    test_hit = txn(counterparty="Acme", amount=-20, day="2026-04-01", tid="hit")
    test_wrong = txn(counterparty="Acme", amount=-21, day="2026-04-02", tid="wrong")
    test_none = txn(counterparty="Corner Cafe", amount=-5, day="2026-04-03", tid="none")
    legacy = [labelled(date(2026, 1, i), -i, GROCERIES, "ACME") for i in range(1, 4)]
    legacy += [labelled(date(2026, 4, 1), -20, GROCERIES, "ACME"), labelled(date(2026, 4, 2), -21, CAFE, "ACME"),
               labelled(date(2026, 4, 3), -5, CAFE, "Corner Cafe")]
    ev = evaluate(train + [test_hit, test_wrong, test_none], legacy, {}, TAXONOMY, [], "common", date(2026, 3, 1))
    assert ev.mined == 1
    assert (ev.overall.total, ev.overall.covered, ev.overall.exact, ev.overall.type_ok) == (3, 2, 1, 1)
    assert ev.uncovered_labels[CAFE] == 1 and len(ev.misses) == 1
    text = render(ev, coverage_by_account(train, [], TAXONOMY, "common"))
    assert "Misses (1)" in text and "Testbank / Alice Current: 0/3" in text
