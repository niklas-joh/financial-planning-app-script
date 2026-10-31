from categoriser.legacy import LegacyRow
from categoriser.mine import (MinedRule, build_examples, existing_signatures, mine_rules, reviewed_labels,
                              _example)
from categoriser.rules import RCOL, parse_rules
from categoriser.schema import CATEGORISED_COLUMNS
from builders import OWNERS, TAXONOMY, rule_row, rules_sheet, txn
from datetime import date

GROCERIES = ("Essentials", "Groceries", "Supermarket")
CAFE = ("Wants/Pleasure", "Eating out", "Cafe")
LUNCH = ("Wants/Pleasure", "Eating out", "Lunch")


def examples(n, label, **kw):
    return [_example(txn(**kw), label) for _ in range(n)]


def test_mines_the_longest_agreeing_counterparty_needle():
    current = [txn(counterparty="Site 12 Acme Market Central")]
    ex = examples(3, GROCERIES, counterparty="ACME MARKET NORTH") + examples(1, CAFE, counterparty="Central Cafe")
    [rule] = mine_rules(current, ex, TAXONOMY)
    assert (rule.field, rule.needle, rule.label, rule.direction) == ("counterpartyContains", "acme market", GROCERIES, "out")
    assert rule.examples == 3 and rule.agreement == 1.0


def test_needs_enough_examples_and_agreement():
    current = [txn(counterparty="Acme")]
    assert mine_rules(current, examples(2, GROCERIES, counterparty="Acme"), TAXONOMY) == []
    mixed = examples(8, GROCERIES, counterparty="Acme") + examples(2, CAFE, counterparty="Acme")
    assert mine_rules(current, mixed, TAXONOMY) == []


def test_falls_back_to_description_then_provider_category():
    current = [txn(counterparty="", description="Monthly lunch card", category="MEAL")]
    ex = examples(3, LUNCH, description="lunch card top up")
    assert mine_rules(current, ex, TAXONOMY)[0].field == "descriptionContains"
    current = [txn(counterparty="", description="DEPOSIT", category="CASHBACK", amount=1)]
    ex = examples(3, ("Income", "Other", "Cashback"), category="CASHBACK", amount=1)
    [rule] = mine_rules(current, ex, TAXONOMY)
    assert (rule.field, rule.needle, rule.direction) == ("providerCategory", "cashback", "in")


def test_skips_labels_outside_the_taxonomy_and_existing_rules():
    current = [txn(counterparty="Acme")]
    assert mine_rules(current, examples(3, ("Gone", "Old", "Label"), counterparty="Acme"), TAXONOMY) == []
    ex = examples(3, GROCERIES, counterparty="Acme")
    assert mine_rules(current, ex, TAXONOMY, {("counterpartyContains", "acme")}) == []


def test_mixed_directions_give_no_direction():
    ex = examples(2, GROCERIES, counterparty="Acme") + examples(1, GROCERIES, counterparty="Acme", amount=4)
    assert mine_rules([txn(counterparty="Acme")], ex, TAXONOMY)[0].direction == ""


def test_mined_rule_row_is_inactive_valid_and_stable():
    rule = MinedRule("counterpartyContains", "acme", "out", GROCERIES, 5, 0.94)
    values = rule.to_values("2026-10-31")
    assert values[RCOL["active"]] is False and values[RCOL["source"]] == "mined"
    assert values[RCOL["notes"]] == "mined 2026-10-31: 5 labelled rows, 94% agree"
    assert rule.rule_id == MinedRule("counterpartyContains", "acme", "out", CAFE, 9, 1.0).rule_id
    parsed = parse_rules(rules_sheet(values), TAXONOMY, OWNERS, "categoriser_rules")
    assert parsed.rules[0].matches(txn(counterparty="Acme", amount=-2)) and parsed.issues == []
    assert rule.as_rule().matches(txn(counterparty="Acme", amount=-2))


def test_existing_signatures_cover_text_and_provider_category():
    sheet = rules_sheet(rule_row("a", counterpartyContains="Acme  Market"), rule_row("b", providerCategory="CASHBACK"))
    assert existing_signatures(sheet) == {("counterpartyContains", "acme market"), ("providerCategory", "cashback")}


def test_examples_prefer_reviewed_over_legacy_and_keep_unjoined_legacy():
    t = txn(counterparty="Acme", amount=-9, day="2026-02-02", tid="t")
    joined = LegacyRow(date(2026, 2, 2), -9, "Old", *CAFE, counterparty="ACME", description="")
    other = LegacyRow(date(2025, 1, 1), -3, "Old", *LUNCH, counterparty="Canteen", description="Card")
    ex = build_examples([t], [joined, other], {"t": GROCERIES})
    assert [e.label for e in ex] == [GROCERIES, LUNCH]
    assert ex[1].counterparty == ex[1].description == "canteen card"


def test_reviewed_labels_reads_only_complete_reviewed_rows():
    sheet = [list(CATEGORISED_COLUMNS),
             ["a", "alice", *GROCERIES, "manual", "", "", True, ""],
             ["b", "alice", *GROCERIES, "rule", "", "R1", False, ""],
             ["c", "alice", "Essentials", "", "", "manual", "", "", True, ""]]
    assert reviewed_labels(sheet) == {"a": GROCERIES}


def test_suggestion_must_cover_its_own_transaction():
    current = [txn(counterparty="Acme", amount=5)]               # incoming
    ex = examples(3, GROCERIES, counterparty="Acme")             # all outgoing
    assert mine_rules(current, ex, TAXONOMY) == []
