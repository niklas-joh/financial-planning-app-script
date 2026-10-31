import json

import pytest

from categoriser.model import (META_FILE, MODEL_FILE, eligible, features, load, model_text, review_check, save,
                               train)
from categoriser.schema import CategorisedRow, Method
from categoriser.taxonomy import Taxonomy
from builders import TAXONOMY, txn

GROCERIES = ("Essentials", "Groceries", "Supermarket")
CAFE = ("Wants/Pleasure", "Eating out", "Cafe")
LUNCH = ("Wants/Pleasure", "Eating out", "Lunch")
NOW = "2026-10-31T06:15:00Z"


def examples():
    out = [(txn(counterparty="Acme Market", description=f"Shop {i}", amount=-20 - i), GROCERIES, 1.0) for i in range(5)]
    out += [(txn(counterparty="Corner Cafe", amount=-3 - i), CAFE, 1.0) for i in range(5)]
    out += [(txn(counterparty="Noodle Bar", amount=-12 - i), LUNCH, 1.0) for i in range(4)]
    return out


@pytest.fixture(scope="module")
def model():
    return train(examples(), TAXONOMY, NOW)


def test_text_drops_numbers_noise_words_and_single_letters():
    t = txn(counterparty="ACME-Market 12", description="WITHDRAW card 4395 12/07/2026 x", reference="REF9")
    assert model_text(t) == "acme market card"
    assert features(txn(amount=15.0)) == {"text": "", "direction": "in", "size": "4"}


def test_card_rows_are_not_for_the_model():
    assert not eligible(txn(classification="card")) and eligible(txn(classification="transfer"))


def test_predicts_known_merchants_with_a_probability(model):
    (label, confidence), = model.predict([txn(counterparty="ACME MARKET", description="Shop 99", amount=-25)])
    assert label == GROCERIES and 0.5 < confidence <= 1
    assert model.meta["examples"] == 14 and model.meta["labels"] == 3


def test_rows_respect_threshold_cards_and_taxonomy(model):
    acme = txn("bob", -25, counterparty="Acme Market", tid="a")
    card = txn("bob", -25, counterparty="Acme Market", tid="c", classification="card")
    rows = model.rows([acme, card], TAXONOMY, 0.5, NOW)
    assert [(r.transactionId, r.owner, r.method, r.ruleId) for r in rows] == [("a", "bob", Method.MODEL, "")]
    assert rows[0].confidence == round(rows[0].confidence, 2)
    assert model.rows([acme], TAXONOMY, 1.0, NOW) == []                      # not confident enough
    assert model.rows([acme], Taxonomy(frozenset({CAFE})), 0.5, NOW) == []   # label left the taxonomy


def test_labels_outside_the_taxonomy_or_too_rare_are_not_learnt():
    odd = [(txn(counterparty="Odd Shop"), ("No", "Such", "Label"), 1.0)] * 3
    rare = [(txn(counterparty="Rare Shop"), LUNCH, 1.0)]
    m = train(examples()[:10] + odd + rare, TAXONOMY, NOW)
    assert set(m.meta["labelCounts"]) == {" / ".join(GROCERIES), " / ".join(CAFE)}
    with pytest.raises(ValueError):
        train(examples()[:5], TAXONOMY, NOW)  # one label only


def test_save_and_load(tmp_path, model):
    assert load(tmp_path) == (None, "missing")
    save(model, tmp_path)
    loaded, status = load(tmp_path)
    assert status == "loaded" and loaded.meta == model.meta
    assert loaded.predict([txn(counterparty="Corner Cafe")])[0][0] == CAFE


def test_other_sklearn_version_or_broken_file_is_skipped(tmp_path, model):
    save(model, tmp_path)
    meta = json.loads((tmp_path / META_FILE).read_text())
    (tmp_path / META_FILE).write_text(json.dumps({**meta, "sklearn": "0.0.1"}))
    assert load(tmp_path) == (None, "incompatible")
    (tmp_path / META_FILE).write_text(json.dumps(meta))
    (tmp_path / MODEL_FILE).write_bytes(b"not a model")
    assert load(tmp_path) == (None, "unreadable")
    (tmp_path / META_FILE).write_text("{")
    assert load(tmp_path) == (None, "unreadable")


def test_review_check_counts_model_rows_that_disagree_with_reviewed_rows():
    row = lambda tid, label: CategorisedRow(tid, "alice", *label, Method.MODEL, 0.95, "", False, NOW)
    rows = [row("a", GROCERIES), row("b", CAFE), row("c", CAFE)]
    assert review_check(rows, {"a": GROCERIES, "b": LUNCH}) == (2, 1)
