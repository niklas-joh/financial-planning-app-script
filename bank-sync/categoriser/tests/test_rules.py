import pytest

from categoriser.rules import INVALID_RULE, contains_words, first_match, normalise, owner_slug, parse_rules
from builders import OWNERS, TAXONOMY, rule_row, rules_sheet, txn


def parse(*rows):
    return parse_rules(rules_sheet(*rows), TAXONOMY, OWNERS, "categoriser_rules")


def test_normalise_and_whole_word_matching():
    assert normalise("Site - 1120_Corner CAFÉ!") == "site 1120 corner café"
    assert contains_words("Site - 1120_Corner Cafe", "corner cafe")
    assert not contains_words("Corner Cafeteria", "corner cafe")
    assert not contains_words("anything", "  ")


def test_owner_slug_matches_bank_sync():
    assert owner_slug(" Ålice Smith ") == "alice-smith"
    assert owner_slug("Common") == "common"


@pytest.mark.parametrize("match,hit,miss", [
    ({"counterpartyContains": "acme market"}, txn(counterparty="ACME MARKET 12"), txn(counterparty="Acme")),
    ({"descriptionContains": "lunch"}, txn(description="Team lunch"), txn(description="Lunchbox")),
    ({"textRegex": r"ref-\d+"}, txn(reference="REF-42"), txn(reference="REF-")),
    ({"providerCategory": "restaurants"}, txn(category="RESTAURANTS"), txn(category="RETAIL")),
    ({"direction": "in"}, txn(amount=5), txn(amount=-5)),
    ({"direction": "out"}, txn(amount=-5), txn(amount=5)),
    ({"amountMin": 10, "amountMax": 20}, txn(amount=-15), txn(amount=-25)),
    ({"account": "testbank / bob current"}, txn(account="bob"), txn(account="alice")),
])
def test_each_match_column(match, hit, miss):
    rule = parse(rule_row(**match)).rules[0]
    assert rule.matches(hit) and not rule.matches(miss)


def test_amount_range_uses_amount_eur_magnitude():
    rule = parse(rule_row(amountMin=10, counterpartyContains="acme")).rules[0]
    assert rule.matches(txn(counterparty="Acme", amount=-99, amountEur=-12))
    assert not rule.matches(txn(counterparty="Acme", amount=-99, amountEur=-8))


def test_all_filled_columns_must_match():
    rule = parse(rule_row(counterpartyContains="acme", direction="out")).rules[0]
    assert rule.matches(txn(counterparty="Acme", amount=-1))
    assert not rule.matches(txn(counterparty="Acme", amount=1))


def test_order_is_priority_then_specificity_then_id():
    parsed = parse(
        rule_row("b-general", counterpartyContains="acme"),
        rule_row("a-general", counterpartyContains="acme"),
        rule_row("specific", counterpartyContains="acme", direction="out"),
        rule_row("urgent", priority=1, descriptionContains="x"),
    )
    assert [r.ruleId for r in parsed.rules] == ["urgent", "specific", "a-general", "b-general"]
    assert first_match(parsed.active, txn(counterparty="Acme", amount=-3)).ruleId == "specific"


def test_inactive_rules_are_parsed_but_not_active():
    parsed = parse(rule_row("off", active=False, counterpartyContains="acme"))
    assert [r.ruleId for r in parsed.rules] == ["off"] and parsed.active == []


@pytest.mark.parametrize("row,reason", [
    (rule_row("", counterpartyContains="a"), "ruleId is empty"),
    (rule_row("R1", priority="soon", counterpartyContains="a"), "priority is not a number"),
    (rule_row("R1", amountMin="ten", counterpartyContains="a"), "amountMin is not a number"),
    (rule_row("R1", amountMin=20, amountMax=10), "amountMin is larger"),
    (rule_row("R1", direction="sideways", counterpartyContains="a"), "direction must be"),
    (rule_row("R1", textRegex="(unclosed"), "textRegex does not compile"),
    (rule_row("R1"), "no match column"),
    (rule_row("R1", output=("Essentials", "Groceries", "Nope"), counterpartyContains="a"), "not in the taxonomy"),
    (rule_row("R1", owner="Carol", counterpartyContains="a"), "owner 'Carol'"),
])
def test_invalid_active_rules_are_skipped_and_reported(row, reason):
    parsed = parse(row, rule_row("ok", counterpartyContains="b"))
    assert [r.ruleId for r in parsed.rules] == ["ok"]
    assert len(parsed.issues) == 1 and parsed.issues[0].type == INVALID_RULE
    assert reason in parsed.issues[0].detail


def test_invalid_inactive_rules_are_skipped_silently():
    parsed = parse(rule_row("R1", active=False, textRegex="("))
    assert parsed.rules == [] and parsed.issues == []


def test_duplicate_rule_ids_are_both_invalid():
    parsed = parse(rule_row("R1", counterpartyContains="a"), rule_row("R1", counterpartyContains="b"))
    assert parsed.rules == [] and len(parsed.issues) == 2
    assert {i.key for i in parsed.issues} == {"invalid_rule:R1"}


def test_issue_key_uses_row_when_rule_id_missing():
    parsed = parse(rule_row("", counterpartyContains="a"))
    assert parsed.issues[0].key == "invalid_rule:row 2"


def test_owner_and_direction_any_are_accepted():
    rule = parse(rule_row("R1", owner="Bob", direction="any", counterpartyContains="a")).rules[0]
    assert rule.owner == "Bob" and rule.direction == ""


def test_blank_rows_are_ignored_and_row_numbers_kept():
    parsed = parse(["" for _ in range(3)], rule_row("R2", counterpartyContains="a"))
    assert parsed.row_of == {"R2": 3}


def test_rejects_a_renamed_header():
    with pytest.raises(ValueError):
        parse_rules([["ruleID"]], TAXONOMY, OWNERS, "categoriser_rules")
