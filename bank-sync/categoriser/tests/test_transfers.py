from categoriser.transfers import OWN_ACCOUNT, PAIR, Transfer, find_transfers, transfer_sub_category
from builders import txn

SUBS = {"Internal Transfer", "From Alice", "From Bob"}


def test_pairs_opposite_amounts_on_two_accounts():
    out, inc = txn("alice", -100, day="2026-05-01"), txn("common", 100, day="2026-05-03")
    found = find_transfers([out, inc])
    assert found[out["transactionId"]] == Transfer(out["transactionId"], PAIR, frozenset({"alice", "common"}))
    assert found[inc["transactionId"]].ruleId == PAIR


def test_does_not_pair_too_far_apart_same_account_other_currency_or_cards():
    assert find_transfers([txn("alice", -100, day="2026-05-01"), txn("common", 100, day="2026-05-05")]) == {}
    assert find_transfers([txn("alice", -100), txn("alice", 100)]) == {}
    assert find_transfers([txn("alice", -100), txn("common", 100, currency="SEK")]) == {}
    assert find_transfers([txn("card", -100, classification="card"), txn("common", 100)]) == {}


def test_pairs_one_to_one_closest_first():
    out = txn("alice", -50, day="2026-05-02")
    near, far = txn("common", 50, day="2026-05-02"), txn("bob", 50, day="2026-05-04")
    found = find_transfers([out, far, near])
    assert near["transactionId"] in found and far["transactionId"] not in found


def test_same_day_same_amount_paired_by_iban_then_description():
    a_out = txn("alice", -700, description="Shared costs A", counterpartyAccount="XX00TEST0003")
    b_out = txn("bob", -700, description="Shared costs B", counterpartyAccount="XX00TEST0003")
    in_b = txn("common", 700, description="Shared costs B", tid="in-b")
    in_a = txn("common", 700, description="Shared costs A", tid="in-a")
    found = find_transfers([b_out, a_out, in_b, in_a])
    assert found["in-a"].owners == {"alice", "common"}
    assert found["in-b"].owners == {"bob", "common"}


def test_single_leg_to_own_iban():
    t = txn("alice", -20, counterpartyAccount="xx00 test 0002")
    other = txn("bob", -1, counterparty="Shop")  # only provides bob's IBAN
    found = find_transfers([t, other])
    assert found[t["transactionId"]] == Transfer(t["transactionId"], OWN_ACCOUNT, frozenset({"alice", "bob"}))
    assert other["transactionId"] not in found


def test_sub_category_names_the_one_person_involved():
    assert transfer_sub_category(Transfer("x", PAIR, frozenset({"alice", "common"})), "common", SUBS) == "From Alice"
    assert transfer_sub_category(Transfer("x", PAIR, frozenset({"bob"})), "common", SUBS) == "From Bob"
    assert transfer_sub_category(Transfer("x", PAIR, frozenset({"alice", "bob"})), "common", SUBS) == "Internal Transfer"
    assert transfer_sub_category(Transfer("x", PAIR, frozenset({"carol", "common"})), "common", SUBS) == "Internal Transfer"
    assert transfer_sub_category(Transfer("x", PAIR, frozenset({"mary-jane"})), "common", {"From Mary-Jane"}) == "From Mary-Jane"
