from categoriser.engine import REVIEW_DISAGREES, SAMPLE_MAX_CHARS, categorise, review_disagreements, sample_text
from categoriser.rules import parse_rules
from categoriser.schema import CATEGORISED_COLUMNS, Method
from builders import OWNERS, TAXONOMY, rule_row, rules_sheet, txn

NOW = "2026-10-31T06:15:00Z"


def rules(*rows):
    return parse_rules(rules_sheet(*rows), TAXONOMY, OWNERS, "categoriser_rules").rules


def by_id(result):
    return {r.transactionId: r for r in result.rows}


def test_rule_row_takes_the_account_owner_unless_the_rule_sets_one():
    a = txn("alice", counterparty="Acme Market")
    b = txn("alice", counterparty="Corner Cafe")
    result = categorise([a, b], rules(
        rule_row("groceries", counterpartyContains="acme market", owner="Common"),
        rule_row("cafe", counterpartyContains="corner cafe", output=("Wants/Pleasure", "Eating out", "Cafe")),
    ), TAXONOMY, "common", NOW)
    rows = by_id(result)
    assert rows[a["transactionId"]].owner == "common" and rows[a["transactionId"]].ruleId == "groceries"
    assert rows[b["transactionId"]].owner == "alice" and rows[b["transactionId"]].subCategory == "Cafe"
    assert rows[a["transactionId"]].method == Method.RULE and rows[a["transactionId"]].confidence is None


def test_uncovered_rows_are_left_out():
    result = categorise([txn(counterparty="Unknown Shop")], rules(rule_row(counterpartyContains="acme")),
                        TAXONOMY, "common", NOW)
    assert result.rows == [] and result.uncovered == 1


def test_inactive_rules_do_not_categorise_but_get_stats():
    t = txn(counterparty="Acme", day="2026-04-02")
    result = categorise([t, txn(counterparty="Acme", day="2026-06-09")],
                        rules(rule_row("off", active=False, counterpartyContains="acme")), TAXONOMY, "common", NOW)
    assert result.rows == []
    assert result.stats["off"].hits == 2 and result.stats["off"].lastMatched == "2026-06-09"


def test_transfer_legs_no_rule_covers_are_labelled():
    out, inc = txn("alice", -300, description="to shared"), txn("common", 300, description="to shared")
    result = categorise([out, inc], [], TAXONOMY, "common", NOW)
    rows = by_id(result)
    assert {r.subCategory for r in rows.values()} == {"From Alice"}
    assert rows[inc["transactionId"]].owner == "common" and rows[out["transactionId"]].ruleId == "transfer:pair"
    assert result.by_transfer == 2


def test_a_rule_wins_over_pairing_and_the_other_leg_still_pairs():
    out = txn("alice", -200, counterparty="Broker Ltd")
    inc = txn("common", 200)
    result = categorise([out, inc], rules(rule_row("save", counterpartyContains="broker ltd",
                                                   output=("Savings", "Savings", "Investments"))),
                        TAXONOMY, "common", NOW)
    rows = by_id(result)
    assert rows[out["transactionId"]].type == "Savings"
    assert rows[inc["transactionId"]].type == "Internal Transfer"


def test_no_transfers_without_transfer_categories_in_taxonomy():
    from categoriser.taxonomy import Taxonomy
    result = categorise([txn("alice", -5), txn("common", 5)], [], Taxonomy(frozenset()), "common", NOW)
    assert result.rows == [] and result.uncovered == 2


def test_duplicate_transaction_ids_are_categorised_once():
    t = txn(counterparty="Acme", tid="same")
    result = categorise([t, dict(t)], rules(rule_row(counterpartyContains="acme")), TAXONOMY, "common", NOW)
    assert [r.transactionId for r in result.rows] == ["same"]


def test_reviewed_rows_that_disagree_with_a_rule_are_reported():
    t, ok = txn(counterparty="Acme", tid="x"), txn(counterparty="Acme", tid="y")
    result = categorise([t, ok], rules(rule_row(counterpartyContains="acme")), TAXONOMY, "common", NOW)
    sheet = [list(CATEGORISED_COLUMNS),
             ["x", "alice", "Wants/Pleasure", "Eating out", "Lunch", "manual", "", "", True, NOW],
             ["y", "alice", "Essentials", "Groceries", "Supermarket", "rule", "", "R1", True, NOW]]
    issues = review_disagreements(sheet, result.rows, "categorised")
    assert [(i.type, i.transactionId) for i in issues] == [(REVIEW_DISAGREES, "x")]
    assert "R1" in issues[0].detail


def test_hits_count_every_match_and_wins_only_the_rows_a_rule_labelled():
    rows = [txn(counterparty="Acme Market"), txn(counterparty="Acme Market Online"), txn(counterparty="Acme Fuel")]
    result = categorise(rows, rules(
        rule_row("market", priority=1, counterpartyContains="acme market"),
        rule_row("acme", priority=2, counterpartyContains="acme"),
        rule_row("off", active=False, counterpartyContains="acme fuel"),
    ), TAXONOMY, "common", NOW)
    got = {rid: (s.hits, s.wins) for rid, s in result.stats.items()}
    assert got == {"market": (2, 2), "acme": (3, 1), "off": (1, 0)}   # inactive: hits, never wins


def test_sample_groups_spellings_and_counts_largest_first():
    rows = ([txn(counterparty="ACME MARKET")] * 2 + [txn(counterparty="Acme-Market")]
            + [txn(counterparty="Corner Cafe")] + [txn(description="card 1234 Fuelstop")]
            + [txn(counterparty="Bakery"), txn(counterparty="Zoo")])
    assert sample_text(rows) == "ACME MARKET (3); Bakery (1); card 1234 Fuelstop (1); +2 more"
    assert sample_text([]) == ""
    long = sample_text([txn(counterparty="x" * 400)])
    assert len(long) == SAMPLE_MAX_CHARS and long.endswith("…")


def test_rule_and_transfer_rows_carry_the_transaction_context():
    g = txn("alice", -12.5, counterparty="Acme Market", description="card payment", day="2026-05-02T10:00:00")
    out, inc = txn("alice", -300, tid="o"), txn("common", 300, tid="i")
    rows = by_id(categorise([g, out, inc], rules(rule_row("R1", counterpartyContains="acme market")),
                            TAXONOMY, "common", NOW))
    assert rows[g["transactionId"]].to_values()[10:] == ["2026-05-02", "Acme Market", "card payment", -12.5]
    assert rows["o"].amountEur == -300 and rows["i"].ruleId == "transfer:pair"
