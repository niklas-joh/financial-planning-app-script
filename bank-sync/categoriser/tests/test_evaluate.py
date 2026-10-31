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


LUNCH = ("Wants/Pleasure", "Eating out", "Lunch")


def test_model_labels_what_the_rules_leave_and_reports_a_curve():
    train = [txn(counterparty="Acme", amount=-i, day=f"2026-01-0{i}") for i in range(1, 4)]
    noodles = txn(counterparty="Noodle Bar", amount=-12, day="2026-04-05", tid="noodles")
    card = txn(counterparty="Noodle Bar", amount=-13, day="2026-04-06", tid="card", classification="card")
    legacy = [labelled(date(2026, 1, i), -i, GROCERIES, "ACME") for i in range(1, 4)]
    legacy += [labelled(date(2026, 2, i), -10 - i, LUNCH, "Noodle Bar") for i in range(1, 3)]  # too few for a rule
    legacy += [labelled(date(2026, 4, 5), -12, LUNCH, "Noodle Bar"), labelled(date(2026, 4, 6), -13, LUNCH, "Noodle Bar")]
    ev = evaluate(train + [noodles, card], legacy, {}, TAXONOMY, [], "common", date(2026, 3, 1), min_confidence=0.3)
    assert ev.mined == 1 and ev.model_examples == 5 and ev.model_skipped_card == 1
    assert (ev.by_method["model"].covered, ev.by_method["model"].exact) == (1, 1)
    assert ev.by_method["none"].total == 1                     # the card row
    assert ev.model_curve["all accounts"][0][1] is True
    text = render(ev, {})
    assert "## Model on test rows the rules leave" in text and "| 0.5 | 1/1, 100% exact" in text
