"""Review inbox (`categoriser_review`): the rows that most need the owner, and their answers.

Each daily run first applies the ticked answers (`apply_answers`): each becomes a
reviewed `manual` row on `categorised`, and an `applyTo` other than "This row only"
also adds an active rule on the merchant text, the description text, or both.
After categorising, the inbox is filled again (`plan_inbox`): one row per merchant
without a label, biggest merchants first, plus a few spot checks of rows the
model labelled. Pure: no I/O (see `sheets.write_review`).
"""

from __future__ import annotations

import hashlib
import random
from collections.abc import Callable
from dataclasses import dataclass, field

from .issues import Issue
from .merge import is_reviewed
from .rules import RCOL, account_label, amount_eur, contains_words, normalise, owner_slug, rule_columns
from .schema import CATEGORISED_COLUMNS, CategorisedRow, Method, transaction_context
from .taxonomy import Taxonomy, Triple

REVIEW_COLUMNS: tuple[str, ...] = (
    # the transaction, and why it is listed (filled by the job)
    "transactionId", "bookingDate", "account", "amountEur", "counterparty", "description", "why",
    # filled by the job, changed by the owner
    "answer", "owner", "applyTo", "merchantText", "descriptionText", "accept",
    # filled by the job only: the model's first guess and its probability, and the rows of this merchant
    "suggested", "confidence", "similar",
)
VCOL = {name: i for i, name in enumerate(REVIEW_COLUMNS)}
# The columns of the inbox before `applyTo` (one `allLikeThis` tick, one `matchText`, a `kind`).
LEGACY_REVIEW_COLUMNS: tuple[str, ...] = (
    "transactionId", "bookingDate", "account", "amountEur", "counterparty", "description",
    "why", "similar", "confidence",
    "answer", "owner", "accept", "allLikeThis", "matchText",
    "suggested", "kind",
)

ANSWER_SEP = " › "  # not "/": "Wants/Pleasure" holds a slash
APPLY_ROW, APPLY_MERCHANT, APPLY_DESCRIPTION, APPLY_BOTH = (
    "This row only", "Same merchant", "Description only", "Merchant + description")
APPLY_TO: tuple[str, ...] = (APPLY_ROW, APPLY_MERCHANT, APPLY_DESCRIPTION, APPLY_BOTH)
REVIEW_RULE_PRIORITY = 50  # ahead of mined rules (default 100)
MIN_MATCH_TEXT = 3  # shorter match text would match far too much
RULE_FIELDS = {"counterpartyContains": "counterparty", "descriptionContains": "description"}
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
    a rule built from them always matches the row (`apply_answers` checks an edited
    `merchantText` / `descriptionText` still does)."""
    words = _longest_plain_run(txn.get("counterparty", ""))
    if words:
        return "counterpartyContains", words, direction(txn)
    return "descriptionContains", _longest_plain_run(txn.get("description", "")), direction(txn)


def review_rule_id(needles: dict[str, str], direction: str) -> str:
    """Stable id of the rule an answer makes: the same match words and direction give the same id.
    (A single needle hashes as it did when a rule had one match text, so earlier rules keep their id.)"""
    key = "|".join(f"{name}|{normalise(words)}" for name, words in needles.items())
    return "r-" + hashlib.sha1(f"{key}|{direction}".encode()).hexdigest()[:8]


def _cell(row: list, name: str) -> object:
    i = VCOL[name]
    return row[i] if len(row) > i else ""


def is_spot_check(row: list) -> bool:
    """A spot check is a row the model had already labelled: unlike a merchant row it has no `similar` count."""
    value = _cell(row, "similar")
    return not (isinstance(value, (int, float)) and not isinstance(value, bool) and value > 0)


OWNER_COLUMNS = ("transactionId", "answer", "owner", "applyTo", "merchantText", "descriptionText", "accept")
_CHECKBOXES = ("accept",)


def owner_cells(review: list[list]) -> list[tuple]:
    """Per inbox row (`review` incl. header): the transactionId and the cells the owner
    edits, ticks as booleans and text stripped, without trailing blank rows. Two reads
    differ exactly when the owner changed something in between."""
    cells = [tuple(is_reviewed(_cell(r, n)) if n in _CHECKBOXES else str(_cell(r, n)).strip()
                   for n in OWNER_COLUMNS) for r in review[1:]]
    blank = tuple(False if n in _CHECKBOXES else "" for n in OWNER_COLUMNS)
    while cells and cells[-1] == blank:
        cells.pop()
    return cells


def transactions_by_id(transactions: list[dict]) -> dict[str, dict]:
    """Transactions by transactionId; the first row wins for a duplicated id (logged elsewhere)."""
    by_id: dict[str, dict] = {}
    for t in transactions:
        tid = str(t.get("transactionId", "")).strip()
        if tid:
            by_id.setdefault(tid, t)
    return by_id


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
            label = [str(row[_CCOL[f]]).strip() for f in _LABEL]
            label[0] = owner_slug(label[0])  # a hand-typed "Bob" is the owner "bob"
            reviewed[tid] = tuple(label)
    return counts, reviewed


def _number(value: object) -> float | None:
    try:
        return float(value) if value not in ("", None) and not isinstance(value, bool) else None
    except (TypeError, ValueError):
        return None


def _rule_values(col: dict[str, int], rule_id: str, needles: dict[str, str], way: str, owner: str,
                 triple: Triple, today: str) -> list:
    """A row for the rules sheet, in the column order `col` (`rules.rule_columns`)."""
    values = [""] * (max(col.values()) + 1)
    named = {
        "ruleId": rule_id, "active": True, "priority": REVIEW_RULE_PRIORITY, "direction": way,
        **needles, "owner": owner, "type": triple[0], "category": triple[1], "subCategory": triple[2],
        "notes": f"from review {today}", "source": "review",
    }
    for name, value in named.items():
        values[col[name]] = value
    return values


def rule_needles(txn: dict, scope: str, merchant_text: str, description_text: str) -> tuple[dict[str, str], str]:
    """(rule match columns -> words, error) for an `applyTo` other than "This row only".

    "Same merchant" matches the merchant text (the description text for a row with no usable
    counterparty); "Description only" matches the description text alone; "Merchant + description"
    needs both. Each text must be at least `MIN_MATCH_TEXT` characters and whole words of this
    row's own counterparty / description, so the rule always catches the row it came from. The error is empty when the needles are valid."""
    uses_counterparty = merchant_key(txn)[0] == "counterpartyContains"
    if scope == APPLY_BOTH:
        wanted = {"counterpartyContains": ("merchantText", merchant_text),
                  "descriptionContains": ("descriptionText", description_text)}
    elif scope == APPLY_MERCHANT and uses_counterparty:
        wanted = {"counterpartyContains": ("merchantText", merchant_text)}
    else:
        wanted = {"descriptionContains": ("descriptionText", description_text)}
    needles: dict[str, str] = {}
    for name, (cell_name, text) in wanted.items():
        words, source = normalise(text), RULE_FIELDS[name]
        if len(words) < MIN_MATCH_TEXT or not contains_words(txn.get(source), words):
            return {}, (f"{cell_name} must be at least {MIN_MATCH_TEXT} characters and whole words of this row's "
                        f"{source}; fix it or set applyTo to '{APPLY_ROW}'.")
        needles[name] = words
    return needles, ""


def apply_answers(review: list[list], transactions: dict[str, dict], taxonomy: Taxonomy, owners: set[str],
                  categorised: list[list], rules: list[list], today: str, now: str, sheet: str) -> AnswerPlan:
    """Turn ticked inbox rows into reviewed manual rows (and rules for an `applyTo` beyond the row).

    `review`, `categorised` and `rules` are whole sheets incl. header; new rules
    follow the rules sheet's column order and skip ruleIds already on it. An answer
    that cannot be applied stays in the inbox (`keep`) with an issue; nothing here
    raises on a bad answer.
    """
    plan = AnswerPlan()
    rule_col = rule_columns(rules[0]) if rules else RCOL
    id_col = rule_col["ruleId"]
    rule_ids = {str(r[id_col]).strip() for r in rules[1:] if len(r) > id_col and str(r[id_col]).strip()}
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
        scope = _text(row, "applyTo") or APPLY_ROW
        if scope not in APPLY_TO:
            problem(row, tid, INVALID_REVIEW, f"applyTo must be one of: {', '.join(APPLY_TO)}.")
            continue
        if scope != APPLY_ROW:
            needles, error = rule_needles(txn, scope, _text(row, "merchantText"), _text(row, "descriptionText"))
            if error:
                problem(row, tid, INVALID_REVIEW, error)
                continue
            way = direction(txn)
            rule_id = review_rule_id(needles, way)
            if rule_id not in rule_ids and rule_id not in added_rules:
                added_rules.add(rule_id)
                account_owner = owner_slug(txn.get("owner", ""))
                plan.new_rules.append(_rule_values(rule_col, rule_id, needles, way,
                                                   "" if owner == account_owner else owner, triple, today))

        plan.manual_rows.append(CategorisedRow(tid, owner, *triple, Method.MANUAL, None, rule_id, True, now,
                                               **transaction_context(txn)))
        plan.processed.add(tid)
        answer, suggested = answer_text(triple), _text(row, "suggested")
        spot = is_spot_check(row)
        plan.changed += bool(suggested) and suggested != answer
        if spot:
            plan.spot_right += suggested == answer
            plan.spot_wrong += suggested != answer
        plan.history.append({"date": today, "transactionId": tid, "suggested": suggested,
                             "confidence": _number(_cell(row, "confidence")), "answer": answer,
                             "spotCheck": spot, "card": str(txn.get("classification", "")).strip().lower() == "card"})
    return plan


Guess = Callable[[list[dict]], list[tuple[Triple, float]]]  # `Model.predict` has this shape


def _padded(row: list) -> list:
    return list(row[: len(REVIEW_COLUMNS)]) + [""] * (len(REVIEW_COLUMNS) - len(row))


def _new_row(txn: dict, why: str, similar: object) -> list:
    values = [""] * len(REVIEW_COLUMNS)
    values[VCOL["transactionId"]] = str(txn.get("transactionId", "")).strip()
    values[VCOL["bookingDate"]] = str(txn.get("bookingDate", ""))[:10]
    values[VCOL["account"]] = account_label(txn)
    values[VCOL["amountEur"]] = amount_eur(txn)
    values[VCOL["counterparty"]] = txn.get("counterparty", "")
    values[VCOL["description"]] = txn.get("description", "")
    values[VCOL["why"]] = why
    values[VCOL["similar"]] = similar
    values[VCOL["owner"]] = owner_slug(txn.get("owner", ""))
    values[VCOL["applyTo"]] = APPLY_ROW
    field_name, words, _way = merchant_key(txn)
    values[VCOL["merchantText"]] = words if field_name == "counterpartyContains" else ""
    values[VCOL["descriptionText"]] = _longest_plain_run(txn.get("description", ""))
    for checkbox in _CHECKBOXES:
        values[VCOL[checkbox]] = False
    return values


def _similar(row: list) -> int:
    value = _cell(row, "similar")
    return int(value) if isinstance(value, (int, float)) and not isinstance(value, bool) else 0


def plan_inbox(previous: list[list], answers: AnswerPlan, by_id: dict[str, dict], labelled: set[str],
               model_ids: set[str], taxonomy: Taxonomy, guess: Guess | None, size: int, spot_checks: int,
               rng: random.Random) -> list[list]:
    """The inbox rows for this run (no header).

    `previous` is the review tab as read at the start of the run, `by_id` the
    transactions (`transactions_by_id`), `labelled` the ids with a label on
    `categorised` after this run's write, `model_ids` the ids the model labelled this run. Rows the owner may still be working on are kept unchanged and
    never dropped; new rows fill up to `size`, the last `spot_checks` of them spot checks.
    """
    kept: list[list] = [_padded(r) for r in answers.keep]
    kept_ids = {_text(r, "transactionId") for r in kept}
    for row in previous[1:]:
        row = _padded(row)
        tid = _text(row, "transactionId")
        if (not tid or tid in kept_ids or tid in answers.processed or tid not in by_id
                or is_reviewed(_cell(row, "accept"))):
            continue
        if (tid in model_ids) if is_spot_check(row) else (tid not in labelled):
            kept.append(row)
            kept_ids.add(tid)

    def group_key(t: dict) -> tuple[str, ...]:
        """A row with no usable merchant words is a group of its own, not one of many unrelated rows."""
        key = merchant_key(t)
        return key if key[1] else (*key, str(t.get("transactionId", "")).strip())

    taken_keys = {group_key(by_id[tid]) for tid in kept_ids if tid in by_id}
    groups: dict[tuple[str, ...], list[dict]] = {}
    for tid, t in by_id.items():
        if tid not in labelled and tid not in kept_ids:
            key = group_key(t)
            if key not in taken_keys:
                groups.setdefault(key, []).append(t)

    kept_spots = sum(is_spot_check(r) for r in kept)
    spot_new = max(0, spot_checks - kept_spots)
    merchant_new = max(0, size - len(kept) - spot_new)

    chosen: list[tuple[dict, str, int]] = []
    for key, members in sorted(groups.items(), key=lambda kv: (-len(kv[1]), kv[0]))[:merchant_new]:
        latest = max(members, key=lambda t: (str(t.get("bookingDate", "")), str(t.get("transactionId", ""))))
        n = len(members)
        why = "No category yet · 1 row like this" if n == 1 else f"No category yet · {n} rows like this"
        chosen.append((latest, why, n))
        taken_keys.add(key)

    spot_pool = sorted(tid for tid in model_ids
                       if tid in by_id and tid not in kept_ids and group_key(by_id[tid]) not in taken_keys)
    spots = rng.sample(spot_pool, min(spot_new, len(spot_pool)))

    new_txns = [t for t, *_ in chosen] + [by_id[tid] for tid in spots]
    guesses = guess(new_txns) if guess is not None and new_txns else [None] * len(new_txns)
    merchant_rows, spot_rows = [], []
    for i, txn in enumerate(new_txns):
        is_merchant = i < len(chosen)
        _t, why, similar = chosen[i] if is_merchant else (txn, "", "")  # a spot check has no `similar`
        row = _new_row(txn, why, similar)
        if guesses[i] is not None and tuple(guesses[i][0]) in taxonomy:
            triple, p = guesses[i]
            row[VCOL["answer"]] = row[VCOL["suggested"]] = answer_text(tuple(triple))
            row[VCOL["confidence"]] = round(p, 2)
        if is_merchant:
            merchant_rows.append(row)
        else:
            confidence = row[VCOL["confidence"]]
            row[VCOL["why"]] = (f"Check the model: {round(confidence * 100)}% sure" if confidence != ""
                                else "Check the model")
            spot_rows.append(row)

    kept_merchants = [r for r in kept if not is_spot_check(r)]
    kept_spot_rows = [r for r in kept if is_spot_check(r)]
    ordered = sorted(kept_merchants + merchant_rows, key=lambda r: -_similar(r))
    return ordered + kept_spot_rows + spot_rows


def is_legacy_review_header(header: list) -> bool:
    return [str(h).strip() for h in header] == list(LEGACY_REVIEW_COLUMNS)


def upgrade_legacy_review(values: list[list]) -> list[list]:
    """The review tab (incl. header) from before `applyTo` in today's columns, keeping every answer.

    A ticked `allLikeThis` becomes "Same merchant" with the old `matchText` as `merchantText`
    (as `descriptionText` for a row with no usable counterparty, as the old rule matched the
    description then). A spot check keeps no `similar`; any other row gets at least 1."""
    old = {name: i for i, name in enumerate(LEGACY_REVIEW_COLUMNS)}

    def get(row: list, name: str) -> object:
        return row[old[name]] if len(row) > old[name] else ""

    upgraded: list[list] = [list(REVIEW_COLUMNS)]
    for row in values[1:]:
        new = [""] * len(REVIEW_COLUMNS)
        for name in REVIEW_COLUMNS:
            if name in old:
                new[VCOL[name]] = get(row, name)
        if any(str(c).strip() for c in row):
            has_counterparty = bool(_longest_plain_run(get(row, "counterparty")))
            match_text = str(get(row, "matchText")).strip()
            new[VCOL["applyTo"]] = APPLY_MERCHANT if is_reviewed(get(row, "allLikeThis")) else APPLY_ROW
            new[VCOL["merchantText"]] = match_text if has_counterparty else ""
            new[VCOL["descriptionText"]] = (_longest_plain_run(get(row, "description")) if has_counterparty
                                            else match_text)
            if str(get(row, "kind")).strip() == "spot check":
                new[VCOL["similar"]] = ""
            elif is_spot_check(new):
                new[VCOL["similar"]] = 1
        upgraded.append(new)
    return upgraded
