from datetime import date

import pytest

from categoriser.legacy import LegacyRow, join_legacy, parse_legacy
from builders import txn

HEADER = ["Source", "Shared", "Type", "Category", "Sub-Category", "Operation", "Date", "Amount", "Currency",
          "Beneficiary Name", "Communication", "Currency", "Category"]


def legacy_row(day, amount, label=("Essentials", "Groceries", "Supermarket"), name="ACME", op="Card payment"):
    return LegacyRow(day, amount, "Old account", *label, counterparty=name, description=op)


def test_parses_labelled_rows_by_header_with_serial_and_iso_dates():
    rows = parse_legacy([HEADER,
                         ["Old", True, "Essentials", "Groceries", "Supermarket", "Card ACME", 46000, -12.5, "EUR", "ACME", "ref", "", "x"],
                         ["Old", "", "Wants/Pleasure", "Eating out", "Cafe", "Cafe", "2026-01-02", "-3", "EUR", "", "", "", ""],
                         ["Old", "", "", "", "", "Unlabelled", 46000, -1, "EUR", "", "", "", ""],
                         ["Old", "", "Essentials", "Groceries", "Supermarket", "No amount", 46000, "", "EUR", "", "", "", ""]])
    assert [(r.day, r.amount, r.label) for r in rows] == [
        (date(2025, 12, 9), -12.5, ("Essentials", "Groceries", "Supermarket")),
        (date(2026, 1, 2), -3.0, ("Wants/Pleasure", "Eating out", "Cafe")),
    ]
    assert rows[0].counterparty == "ACME" and rows[0].description == "Card ACME ref"


def test_missing_columns_are_an_error():
    with pytest.raises(ValueError):
        parse_legacy([["Source", "Type"]])


def test_join_is_one_to_one_by_amount_and_date_window():
    t1 = txn(amount=-12.5, day="2026-01-03", tid="t1")
    t2 = txn(amount=-12.5, day="2026-01-20", tid="t2")
    far, near = legacy_row(date(2026, 1, 1), -12.5), legacy_row(date(2026, 1, 2), -12.5)
    joined = join_legacy([far, near], [t1, t2])
    assert joined == {"t1": near}  # closest date wins; t2 is 18 days off; `far` stays unused


def test_join_breaks_same_day_ties_by_shared_words():
    a = txn(amount=-700, counterparty="ALICE", description="shared costs", tid="a")
    b = txn(amount=-700, counterparty="BOB", description="shared costs", tid="b")
    la = legacy_row(date(2026, 5, 4), -700, name="ALICE", op="shared costs")
    lb = legacy_row(date(2026, 5, 4), -700, name="BOB", op="shared costs")
    joined = join_legacy([lb, la], [b, a])
    assert joined["a"] is la and joined["b"] is lb
