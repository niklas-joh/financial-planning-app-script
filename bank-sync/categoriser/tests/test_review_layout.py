import hashlib

from categoriser.review import (
    answer_text,
    is_spot_check,
    merchant_key,
    owner_cells,
    parse_answer,
    review_rule_id,
)
from categoriser.rules import contains_words
from builders import review_row, review_sheet, txn


def test_answer_round_trips_and_keeps_slashes():
    t = ("Wants/Pleasure", "Eating out", "Cafe")
    assert answer_text(t) == "Wants/Pleasure › Eating out › Cafe"
    assert parse_answer("  Wants/Pleasure ›Eating out ›  Cafe ") == t
    assert parse_answer("") is None and parse_answer(None) is None
    assert parse_answer("Essentials › Groceries") is None


def test_merchant_key_is_the_longest_run_of_words_without_numbers():
    assert merchant_key(txn(counterparty="Corner_Cafe 1234 Ltd", amount=-3)) == ("counterpartyContains", "corner cafe", "out")
    assert merchant_key(txn(counterparty="PAYPAL *ACME 4029357733")) == ("counterpartyContains", "paypal acme", "out")


def test_merchant_key_falls_back_to_the_description():
    assert merchant_key(txn(counterparty="12345", description="SEPA Rent 2026", amount=900)) == \
        ("descriptionContains", "sepa rent", "in")
    assert merchant_key(txn(counterparty="", description="")) == ("descriptionContains", "", "out")


def test_merchant_key_words_always_match_their_row():
    t = txn(counterparty="Shop 12 - Corner_Cafe Ltd")
    assert contains_words(t["counterparty"], merchant_key(t)[1])


def test_review_rule_id_is_stable_and_normalised():
    a = review_rule_id({"counterpartyContains": "Corner-Cafe"}, "out")
    assert a == review_rule_id({"counterpartyContains": "corner cafe"}, "out") and a.startswith("r-") and len(a) == 10
    assert a != review_rule_id({"counterpartyContains": "corner cafe"}, "in")


def test_review_rule_id_keeps_the_id_of_a_rule_made_with_one_match_text():
    # sha1("counterpartyContains|corner cafe|out") before rules could have two needles
    assert review_rule_id({"counterpartyContains": "corner cafe"}, "out") == "r-" + hashlib.sha1(
        b"counterpartyContains|corner cafe|out").hexdigest()[:8]


def test_review_rule_id_depends_on_every_needle():
    one = review_rule_id({"counterpartyContains": "corner cafe"}, "out")
    two = review_rule_id({"counterpartyContains": "corner cafe", "descriptionContains": "gift"}, "out")
    other = review_rule_id({"counterpartyContains": "corner cafe", "descriptionContains": "tip"}, "out")
    assert len({one, two, other}) == 3


def test_owner_cells_read_ticks_as_booleans_and_ignore_trailing_blank_rows():
    as_read = review_sheet(review_row("a", accept=True, answer=" A › B › C"), review_row("b", accept="FALSE"))
    fresh = review_sheet(review_row("a", accept="TRUE", answer="A › B › C"), review_row("b", accept=""), [""] * 3)
    assert owner_cells(as_read) == owner_cells(fresh) and owner_cells([]) == []
    assert owner_cells(as_read)[0] == ("a", "A › B › C", "", "This row only", "", "", True)


def test_owner_cells_see_each_owner_column_and_the_job_columns_do_not_count():
    base = review_sheet(review_row("a", merchant_text="acme"))
    for edit in (dict(answer="A › B › C"), dict(owner="bob"), dict(accept=True), dict(apply_to="Same merchant"),
                 dict(merchant_text="acme market"), dict(description_text="gift"), dict(tid="b")):
        fields = dict(tid="a", merchant_text="acme") | edit
        assert owner_cells(review_sheet(review_row(**fields))) != owner_cells(base), edit
    assert owner_cells(review_sheet(review_row("a", merchant_text="acme", suggested="X › Y › Z"))) == owner_cells(base)


def test_a_row_is_a_spot_check_when_it_has_no_similar_count():
    assert is_spot_check(review_row("a", spot_check=True)) and not is_spot_check(review_row("a"))
    assert is_spot_check([]) and is_spot_check(["x"] * 16)
