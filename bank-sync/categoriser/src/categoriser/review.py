"""Review inbox (`categoriser_review`): the rows that most need the owner, and their answers.

Each daily run first applies the ticked answers (`apply_answers`): each becomes a
reviewed `manual` row on `categorised`, and "all like this" also adds an active
rule. After categorising, the inbox is filled again (`plan_inbox`): one row per
merchant without a label, biggest merchants first, plus a few spot checks of
rows the model labelled. Pure: no I/O (see `sheets.write_review`).
"""

from __future__ import annotations

import hashlib
from dataclasses import dataclass, field

from .issues import Issue
from .merge import is_reviewed
from .rules import RCOL, RULE_COLUMNS, amount_eur, normalise, owner_slug
from .schema import CATEGORISED_COLUMNS, CategorisedRow, Method
from .taxonomy import Taxonomy, Triple

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


def _text(row: list, name: str) -> str:
    return str(_cell(row, name)).strip()


@dataclass
class AnswerPlan:
    manual_rows: list[CategorisedRow] = field(default_factory=list)  # reviewed=True, method=manual
    new_rules: list[list] = field(default_factory=list)              # rows for categoriser_rules
    keep: list[list] = field(default_factory=list)                   # ticked inbox rows not applied, as read
    processed: set[str] = field(default_factory=set)                 # answers done: leave the inbox
    issues: list[Issue] = field(default_factory=list)
    history: list[dict] = field(default_factory=list)                # one per manual row written
    changed: int = 0          # manual rows whose answer differs from a non-blank `suggested`
    spot_right: int = 0
    spot_wrong: int = 0


_CCOL = {name: i for i, name in enumerate(CATEGORISED_COLUMNS)}
_LABEL = ("owner", "type", "category", "subCategory")


def _categorised_index(categorised: list[list]) -> tuple[dict[str, int], dict[str, tuple[str, ...]]]:
    """(rows per transactionId, reviewed (owner, type, category, subCategory) per transactionId)."""
    counts: dict[str, int] = {}
    reviewed: dict[str, tuple[str, ...]] = {}
    for row in categorised[1:]:
        row = list(row) + [""] * (len(CATEGORISED_COLUMNS) - len(row))
        tid = str(row[_CCOL["transactionId"]]).strip()
        if not tid:
            continue
        counts[tid] = counts.get(tid, 0) + 1
        if is_reviewed(row[_CCOL["reviewed"]]):
            reviewed[tid] = tuple(str(row[_CCOL[f]]).strip() for f in _LABEL)
    return counts, reviewed


def _number(value: object) -> float | None:
    try:
        return float(value) if value not in ("", None) and not isinstance(value, bool) else None
    except (TypeError, ValueError):
        return None


def _rule_values(rule_id: str, field_name: str, words: str, way: str, owner: str, triple: Triple,
                 today: str) -> list:
    values = [""] * len(RULE_COLUMNS)
    values[RCOL["ruleId"]] = rule_id
    values[RCOL["active"]] = True
    values[RCOL["priority"]] = REVIEW_RULE_PRIORITY
    values[RCOL["direction"]] = way
    values[RCOL[field_name]] = words
    values[RCOL["owner"]] = owner
    values[RCOL["type"]], values[RCOL["category"]], values[RCOL["subCategory"]] = triple
    values[RCOL["notes"]] = f"from review {today}"
    values[RCOL["source"]] = "review"
    return values


def apply_answers(review: list[list], transactions: dict[str, dict], taxonomy: Taxonomy, owners: set[str],
                  categorised: list[list], rule_ids: set[str], today: str, now: str, sheet: str) -> AnswerPlan:
    """Turn ticked inbox rows into reviewed manual rows (and rules for "all like this").

    `review` and `categorised` are whole sheets incl. header; `rule_ids` are the
    ruleIds already on the rules sheet. An answer that cannot be applied stays in
    the inbox (`keep`) with an issue; nothing here raises on a bad answer.
    """
    plan = AnswerPlan()
    counts, reviewed = _categorised_index(categorised)
    added_rules: set[str] = set()
    handled: set[str] = set()
    source = f"categoriser.run: '{sheet}' answers"

    def problem(row: list, tid: str, kind: str, detail: str) -> None:
        plan.keep.append(row)
        plan.issues.append(Issue(type=kind, transactionId=tid, detail=detail, source=source))

    for row in review[1:]:
        row = list(row) + [""] * (len(REVIEW_COLUMNS) - len(row))
        tid = _text(row, "transactionId")
        if not tid or not is_reviewed(_cell(row, "accept")) or tid in handled:
            continue
        handled.add(tid)
        txn = transactions.get(tid)
        triple = parse_answer(_cell(row, "answer"))
        owner = owner_slug(_text(row, "owner")) or owner_slug((txn or {}).get("owner", ""))
        if txn is None:
            problem(row, tid, INVALID_REVIEW, f"Not on transactions any more; delete the row in '{sheet}'.")
            continue
        if triple is None or triple not in taxonomy:
            problem(row, tid, INVALID_REVIEW, "The answer is not a type › category › subCategory of the taxonomy.")
            continue
        if owner not in owners:
            problem(row, tid, INVALID_REVIEW, f"Unknown owner '{owner}'.")
            continue
        if counts.get(tid, 0) > 1:
            problem(row, tid, INVALID_REVIEW, "On several rows of categorised; delete the extra rows first.")
            continue
        if tid in reviewed:
            if reviewed[tid] != (owner, *triple):
                problem(row, tid, REVIEW_CONFLICT,
                        f"categorised already has a reviewed row with {' / '.join(reviewed[tid])}. Change that "
                        f"row, or untick or delete this one in '{sheet}'.")
            else:
                plan.processed.add(tid)
            continue

        rule_id = ""
        words = normalise(_cell(row, "matchText"))
        if is_reviewed(_cell(row, "allLikeThis")) and words:
            field_name, _key_words, way = merchant_key(txn)
            rule_id = review_rule_id(field_name, words, way)
            if rule_id not in rule_ids and rule_id not in added_rules:
                added_rules.add(rule_id)
                account_owner = owner_slug(txn.get("owner", ""))
                plan.new_rules.append(_rule_values(rule_id, field_name, words, way,
                                                   "" if owner == account_owner else owner, triple, today))

        plan.manual_rows.append(CategorisedRow(tid, owner, *triple, Method.MANUAL, None, rule_id, True, now))
        plan.processed.add(tid)
        answer, suggested = answer_text(triple), _text(row, "suggested")
        spot = _text(row, "kind") == KIND_SPOT_CHECK
        plan.changed += bool(suggested) and suggested != answer
        if spot:
            plan.spot_right += suggested == answer
            plan.spot_wrong += suggested != answer
        plan.history.append({"date": today, "transactionId": tid, "suggested": suggested,
                             "confidence": _number(_cell(row, "confidence")), "answer": answer,
                             "spotCheck": spot, "card": str(txn.get("classification", "")).strip().lower() == "card"})
    return plan
