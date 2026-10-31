"""Review inbox (`categoriser_review`): the rows that most need the owner, and their answers.

Each daily run first applies the ticked answers (`apply_answers`): each becomes a
reviewed `manual` row on `categorised`, and "all like this" also adds an active
rule. After categorising, the inbox is filled again (`plan_inbox`): one row per
merchant without a label, biggest merchants first, plus a few spot checks of
rows the model labelled. Pure: no I/O (see `sheets.write_review`).
"""

from __future__ import annotations

import hashlib

from .merge import is_reviewed
from .rules import amount_eur, normalise
from .taxonomy import Triple

REVIEW_COLUMNS: tuple[str, ...] = (
    "transactionId", "bookingDate", "account", "amountEur", "counterparty", "description",
    "why", "similar", "confidence",
    # filled by the job, changed by the owner
    "answer", "owner", "accept", "allLikeThis", "matchText",
    # filled by the job only
    "suggested", "kind",
)
VCOL = {name: i for i, name in enumerate(REVIEW_COLUMNS)}

ANSWER_SEP = " › "  # not "/": "Wants/Pleasure" holds a slash
KIND_MERCHANT, KIND_SPOT_CHECK = "merchant", "spot check"
REVIEW_RULE_PRIORITY = 50  # ahead of mined rules (default 100)
INVALID_REVIEW = "invalid_review_answer"
REVIEW_CONFLICT = "review_conflicts_with_reviewed"


def answer_text(triple: Triple) -> str:
    return ANSWER_SEP.join(triple)


def parse_answer(text: object) -> Triple | None:
    """`type › category › subCategory`, or None unless there are exactly three non-empty parts."""
    parts = [p.strip() for p in str(text or "").split(ANSWER_SEP.strip())]
    return (parts[0], parts[1], parts[2]) if len(parts) == 3 and all(parts) else None


def _longest_plain_run(text: object) -> str:
    """The longest run of consecutive words without digits (first on a tie), normalised."""
    best: list[str] = []
    run: list[str] = []
    for word in normalise(text).split() + ["0"]:  # the sentinel closes the last run
        if any(c.isdigit() for c in word):
            if len(run) > len(best):
                best = run
            run = []
        else:
            run.append(word)
    return " ".join(best)


def direction(txn: dict) -> str:
    return "in" if amount_eur(txn) > 0 else "out"


def merchant_key(txn: dict) -> tuple[str, str, str]:
    """(rule field, words, direction) shared by rows of the same merchant. The words are
    the counterparty's longest run of words without numbers, else the description's, so
    a rule built from them always matches the row."""
    words = _longest_plain_run(txn.get("counterparty", ""))
    if words:
        return "counterpartyContains", words, direction(txn)
    return "descriptionContains", _longest_plain_run(txn.get("description", "")), direction(txn)


def review_rule_id(field: str, words: str, direction: str) -> str:
    return "r-" + hashlib.sha1(f"{field}|{normalise(words)}|{direction}".encode()).hexdigest()[:8]


def _cell(row: list, name: str) -> object:
    i = VCOL[name]
    return row[i] if len(row) > i else ""


def ticked(review: list[list]) -> set[str]:
    """transactionIds of inbox rows with `accept` ticked (`review` incl. header)."""
    return {str(_cell(r, "transactionId")).strip() for r in review[1:]
            if is_reviewed(_cell(r, "accept")) and str(_cell(r, "transactionId")).strip()}
