from categoriser.review import answer_text, merchant_key, owner_cells, parse_answer, review_rule_id
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
    a = review_rule_id("counterpartyContains", "Corner-Cafe", "out")
    assert a == review_rule_id("counterpartyContains", "corner cafe", "out") and a.startswith("r-") and len(a) == 10
    assert a != review_rule_id("counterpartyContains", "corner cafe", "in")


def test_owner_cells_read_ticks_as_booleans_and_ignore_trailing_blank_rows():
    as_read = review_sheet(review_row("a", accept=True, answer=" A › B › C"), review_row("b", accept="FALSE"))
    fresh = review_sheet(review_row("a", accept="TRUE", answer="A › B › C"), review_row("b", accept=""), [""] * 3)
    assert owner_cells(as_read) == owner_cells(fresh) and owner_cells([]) == []
    assert owner_cells(as_read)[0] == ("a", "A › B › C", "", True, False, "")


def test_owner_cells_see_each_owner_column_and_the_job_columns_do_not_count():
    base = review_sheet(review_row("a", match_text="acme"))
    for edit in (dict(answer="A › B › C"), dict(owner="bob"), dict(accept=True), dict(all_like_this=True),
                 dict(match_text="acme market"), dict(tid="b")):
        fields = dict(tid="a", match_text="acme") | edit
        assert owner_cells(review_sheet(review_row(**fields))) != owner_cells(base), edit
    assert owner_cells(review_sheet(review_row("a", match_text="acme", suggested="X › Y › Z"))) == owner_cells(base)
