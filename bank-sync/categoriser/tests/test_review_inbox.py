import random

from categoriser.review import REVIEW_COLUMNS, VCOL, AnswerPlan, plan_inbox
from builders import TAXONOMY, review_row, review_sheet, txn

GROCERY = ("Essentials", "Groceries", "Supermarket")


def inbox(txns, previous=None, answers=None, labelled=(), model_ids=(), guess=None, size=25, spot=0):
    return plan_inbox(previous or review_sheet(), answers or AnswerPlan(), txns, set(labelled), set(model_ids),
                      TAXONOMY, guess, size, spot, random.Random(1))


def ids(rows):
    return [r[VCOL["transactionId"]] for r in rows]


def test_biggest_merchant_groups_first_one_latest_row_each():
    txns = [txn(counterparty="Acme Market", day=f"2026-05-0{d}", tid=f"a{d}") for d in (1, 2, 3)]
    txns += [txn(counterparty="Corner Cafe", day=f"2026-05-0{d}", tid=f"c{d}") for d in (1, 2)]
    txns += [txn(counterparty="Lone Shop", tid="l")]
    rows = inbox(txns, size=2)
    assert ids(rows) == ["a3", "c2"]
    assert rows[0][VCOL["why"]] == "3 rows like this have no category" and rows[0][VCOL["similar"]] == 3
    assert rows[0][VCOL["matchText"]] == "acme market" and rows[0][VCOL["kind"]] == "merchant"
    assert all(len(r) == len(REVIEW_COLUMNS) for r in rows)


def test_one_row_group_says_row_in_the_singular():
    [row] = inbox([txn(counterparty="Lone Shop", tid="l")])
    assert row[VCOL["why"]] == "1 row like this has no category"


def test_row_fields_come_from_the_transaction():
    [row] = inbox([txn("bob", -12.5, counterparty="Acme Market", description="Card 1", day="2026-05-04T10:00:00Z",
                       tid="a")])
    assert row[:6] == ["a", "2026-05-04", "Testbank / Bob Current", -12.5, "Acme Market", "Card 1"]


def test_labelled_rows_are_not_candidates():
    assert inbox([txn(counterparty="Acme Market", tid="a")], labelled={"a"}) == []


def test_unanswered_row_survives_with_its_typed_answer_and_its_group_is_not_added_again():
    old = review_row("a1", answer="Essentials › Groceries › Supermarket")
    txns = [txn(counterparty="Acme Market", tid="a1"), txn(counterparty="Acme Market", tid="a2")]
    rows = inbox(txns, previous=review_sheet(old))
    assert rows == [old]


def test_rows_that_got_a_label_leave_the_inbox():
    rows = inbox([txn(counterparty="Acme Market", tid="a1")], previous=review_sheet(review_row("a1")), labelled={"a1"})
    assert rows == []


def test_rows_no_longer_on_transactions_leave_the_inbox():
    assert inbox([], previous=review_sheet(review_row("gone"))) == []


def test_answered_rows_leave_the_inbox():
    done = review_row("a1", answer="Essentials › Groceries › Supermarket", accept=True)
    rows = inbox([txn(counterparty="Acme Market", tid="a1")], previous=review_sheet(done),
                 answers=AnswerPlan(processed={"a1"}), labelled={"a1"})
    assert rows == []


def test_ticked_rows_that_were_not_applied_stay():
    bad = review_row("a1", answer="nope", accept=True)
    rows = inbox([txn(tid="a1")], previous=review_sheet(bad), answers=AnswerPlan(keep=[bad]), labelled={"a1"})
    assert ids(rows) == ["a1"]


def test_spot_checks_take_the_last_slots_from_model_rows():
    txns = [txn(counterparty="Acme Market", tid="a"), txn(counterparty="Corner Cafe", tid="m1"),
            txn(counterparty="Lone Shop", tid="m2")]
    rows = inbox(txns, labelled={"m1", "m2"}, model_ids={"m1", "m2"}, size=2, spot=1,
                 guess=lambda ts: [(GROCERY, 0.96)] * len(ts))
    assert [r[VCOL["kind"]] for r in rows] == ["merchant", "spot check"]
    assert rows[1][VCOL["why"]] == "spot check: model 96% sure" and rows[1][VCOL["transactionId"]] in {"m1", "m2"}
    assert rows[1][VCOL["similar"]] == ""


def test_spot_checks_skip_merchants_already_in_the_inbox():
    txns = [txn(counterparty="Acme Market", tid="a"), txn(counterparty="Acme Market", tid="m")]
    rows = inbox(txns, labelled={"m"}, model_ids={"m"}, size=3, spot=1)
    assert ids(rows) == ["a"]


def test_unanswered_spot_check_stays_while_the_model_still_labels_it():
    old = review_row("m", kind="spot check")
    rows = inbox([txn(counterparty="Corner Cafe", tid="m")], previous=review_sheet(old), labelled={"m"},
                 model_ids={"m"}, spot=1)
    assert rows == [old]


def test_prefill_uses_the_model_guess_only_inside_the_taxonomy():
    txns = [txn(counterparty="Acme Market", tid="a", classification="card"), txn(counterparty="Odd", tid="o")]
    guesses = {"a": (GROCERY, 0.734), "o": (("No", "Such", "Label"), 0.99)}
    rows = {r[0]: r for r in inbox(txns, guess=lambda ts: [guesses[t["transactionId"]] for t in ts])}
    assert rows["a"][VCOL["answer"]] == rows["a"][VCOL["suggested"]] == "Essentials › Groceries › Supermarket"
    assert rows["a"][VCOL["confidence"]] == 0.73
    assert rows["o"][VCOL["answer"]] == rows["o"][VCOL["suggested"]] == rows["o"][VCOL["confidence"]] == ""


def test_guess_is_called_once_for_all_new_rows():
    calls = []
    inbox([txn(counterparty="Acme Market", tid="a"), txn(counterparty="Corner Cafe", tid="c")],
          guess=lambda ts: calls.append(len(ts)) or [(GROCERY, 0.5)] * len(ts))
    assert calls == [2]


def test_without_a_model_the_answer_is_blank_and_owner_is_prefilled():
    [row] = inbox([txn("bob", counterparty="Acme Market", tid="a")])
    assert row[VCOL["answer"]] == "" and row[VCOL["owner"]] == "bob"
    assert row[VCOL["accept"]] is False and row[VCOL["allLikeThis"]] is False


def test_kept_rows_are_never_dropped_and_no_new_rows_beyond_size():
    olds = [review_row(f"k{i}") for i in range(3)]
    txns = [txn(counterparty=f"Shop {chr(97 + i)}", tid=f"k{i}") for i in range(3)] + [txn(counterparty="New", tid="n")]
    assert len(inbox(txns, previous=review_sheet(*olds), size=2)) == 3


def test_kept_and_new_rows_sort_by_similar():
    old = review_row("l")  # similar blank counts as 0
    txns = [txn(counterparty="Lone Shop", tid="l")] + [txn(counterparty="Acme Market", tid=f"a{i}") for i in range(2)]
    assert ids(inbox(txns, previous=review_sheet(old))) == ["a1", "l"]


def test_rows_without_merchant_words_are_not_grouped_together():
    txns = [txn(counterparty="123", description="", tid=f"x{i}") for i in range(3)]
    rows = inbox(txns)
    assert len(rows) == 3 and {r[VCOL["why"]] for r in rows} == {"1 row like this has no category"}
