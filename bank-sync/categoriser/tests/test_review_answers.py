import pytest

from categoriser.review import apply_answers, review_rule_id
from categoriser.rules import RCOL, parse_rules
from categoriser.schema import CATEGORISED_COLUMNS, Method
from builders import OWNERS, TAXONOMY, review_row, review_sheet, rules_sheet, txn

HEADER = list(CATEGORISED_COLUMNS)
GROCERY = "Essentials › Groceries › Supermarket"
CAFE = "Wants/Pleasure › Eating out › Cafe"


def plan(review, txns, categorised=None, rule_ids=frozenset()):
    return apply_answers(review, {t["transactionId"]: t for t in txns}, TAXONOMY, OWNERS,
                         categorised or [HEADER], set(rule_ids), "2026-11-01", "2026-11-01T06:15:00Z",
                         "categoriser_review")


def test_accepted_answer_becomes_a_reviewed_manual_row():
    p = plan(review_sheet(review_row("a", answer=GROCERY, accept=True)), [txn("bob", tid="a")])
    [row] = p.manual_rows
    assert (row.transactionId, row.owner, row.type, row.method, row.reviewed, row.confidence, row.ruleId) == \
        ("a", "bob", "Essentials", Method.MANUAL, True, None, "")
    assert (row.category, row.subCategory, row.updatedAt) == ("Groceries", "Supermarket", "2026-11-01T06:15:00Z")
    assert p.processed == {"a"} and p.keep == [] and p.issues == []


def test_unticked_rows_are_ignored():
    p = plan(review_sheet(review_row("a", answer=GROCERY)), [txn(tid="a")])
    assert p.manual_rows == [] and p.processed == set() and p.keep == []


@pytest.mark.parametrize("answer,owner,tid", [
    ("Essentials › Groceries › Bakery", "", "a"),   # not in the taxonomy
    ("", "", "a"),                                   # no answer
    (GROCERY, "carol", "a"),                         # unknown owner
    (GROCERY, "", "gone"),                           # not on transactions
])
def test_invalid_answers_stay_ticked_with_an_issue(answer, owner, tid):
    row = review_row(tid, answer=answer, owner=owner, accept=True)
    p = plan(review_sheet(row), [txn(tid="a")])
    assert p.manual_rows == [] and p.keep == [row] and p.processed == set()
    assert [(i.type, i.transactionId) for i in p.issues] == [("invalid_review_answer", tid)]


def test_id_on_several_categorised_rows_is_not_applied():
    dup = ["a", "", "", "", "", "", "", "", False, ""]
    p = plan(review_sheet(review_row("a", answer=GROCERY, accept=True)), [txn(tid="a")], [HEADER, dup, dup])
    assert p.manual_rows == [] and [i.type for i in p.issues] == ["invalid_review_answer"]


def test_answer_matching_a_reviewed_row_is_done_without_writing():
    reviewed = ["a", "alice", "Essentials", "Groceries", "Supermarket", "manual", "", "", True, ""]
    p = plan(review_sheet(review_row("a", answer=GROCERY, accept=True)), [txn("alice", tid="a")], [HEADER, reviewed])
    assert p.manual_rows == [] and p.history == [] and p.processed == {"a"} and p.issues == []


def test_answer_conflicting_with_a_reviewed_row_stays_with_an_issue():
    reviewed = ["a", "alice", "Wants/Pleasure", "Eating out", "Cafe", "manual", "", "", True, ""]
    row = review_row("a", answer=GROCERY, accept=True)
    p = plan(review_sheet(row), [txn("alice", tid="a")], [HEADER, reviewed])
    assert p.keep == [row] and p.manual_rows == [] and p.processed == set()
    assert [i.type for i in p.issues] == ["review_conflicts_with_reviewed"]


def test_all_like_this_adds_one_valid_active_rule():
    t1, t2 = txn("bob", counterparty="Acme Market 77", tid="a"), txn("bob", counterparty="Acme Market", tid="b")
    p = plan(review_sheet(review_row("a", answer=GROCERY, accept=True, all_like_this=True, match_text="acme market"),
                          review_row("b", answer=GROCERY, accept=True, all_like_this=True, match_text="acme market")),
             [t1, t2])
    [rule_values] = p.new_rules
    parsed = parse_rules(rules_sheet(rule_values), TAXONOMY, OWNERS, "categoriser_rules")
    [rule] = parsed.active
    assert parsed.issues == [] and rule.matches(t1) and rule.matches(t2)
    assert (rule.priority, rule.source, rule.direction, rule.owner) == (50, "review", "out", "")
    assert rule_values[RCOL["notes"]] == "from review 2026-11-01"
    assert {r.ruleId for r in p.manual_rows} == {rule.ruleId}


def test_existing_rule_id_is_not_added_again():
    rid = review_rule_id("counterpartyContains", "acme market", "out")
    p = plan(review_sheet(review_row("a", answer=GROCERY, accept=True, all_like_this=True, match_text="acme market")),
             [txn(counterparty="Acme Market", tid="a")], rule_ids={rid})
    assert p.new_rules == [] and p.manual_rows[0].ruleId == rid


def test_all_like_this_without_match_text_adds_no_rule():
    p = plan(review_sheet(review_row("a", answer=GROCERY, accept=True, all_like_this=True, match_text=" ")),
             [txn(counterparty="Acme Market", tid="a")])
    assert p.new_rules == [] and p.manual_rows[0].ruleId == ""


def test_rule_keeps_an_owner_other_than_the_accounts():
    p = plan(review_sheet(review_row("a", answer=GROCERY, owner="common", accept=True, all_like_this=True,
                                     match_text="acme market")), [txn("bob", counterparty="Acme Market", tid="a")])
    assert p.new_rules[0][RCOL["owner"]] == "common" and p.manual_rows[0].owner == "common"


def test_history_and_counts():
    p = plan(review_sheet(review_row("a", answer=GROCERY, suggested=CAFE, accept=True),
                          review_row("b", answer=CAFE, suggested=CAFE, accept=True, kind="spot check", confidence=0.96),
                          review_row("c", answer=GROCERY, suggested=CAFE, accept=True, kind="spot check")),
             [txn(tid="a", classification="card"), txn(tid="b"), txn(tid="c")])
    assert (p.changed, p.spot_right, p.spot_wrong) == (2, 1, 1)
    assert p.history[0] == {"date": "2026-11-01", "transactionId": "a", "suggested": CAFE, "confidence": None,
                            "answer": GROCERY, "spotCheck": False, "card": True}
    assert p.history[1]["confidence"] == 0.96 and p.history[1]["spotCheck"] is True


def test_row_order_does_not_matter():
    rows = [review_row("a", answer=GROCERY, accept=True), review_row("b", answer=CAFE, accept=True)]
    txns = [txn(tid="a"), txn(tid="b")]
    forward, backward = plan(review_sheet(*rows), txns), plan(review_sheet(*reversed(rows)), txns)
    assert sorted(forward.manual_rows, key=lambda r: r.transactionId) == \
        sorted(backward.manual_rows, key=lambda r: r.transactionId)


def test_duplicate_inbox_rows_apply_once():
    row = review_row("a", answer=GROCERY, accept=True)
    p = plan(review_sheet(row, list(row)), [txn(tid="a")])
    assert len(p.manual_rows) == 1 and len(p.history) == 1 and p.issues == [] and p.keep == []


def test_ticks_as_text_count():
    p = plan(review_sheet(review_row("a", answer=GROCERY, accept="TRUE"), review_row("b", answer=GROCERY, accept="FALSE"),
                          review_row("c", answer=GROCERY, accept="")), [txn(tid="a"), txn(tid="b"), txn(tid="c")])
    assert [r.transactionId for r in p.manual_rows] == ["a"]


def test_match_text_is_normalised():
    p = plan(review_sheet(review_row("a", answer=CAFE, accept=True, all_like_this=True, match_text="Corner-Cafe")),
             [txn(counterparty="Corner Cafe", tid="a")])
    [rule] = p.new_rules
    assert rule[RCOL["counterpartyContains"]] == "corner cafe"
    assert rule[RCOL["ruleId"]] == review_rule_id("counterpartyContains", "corner cafe", "out")
