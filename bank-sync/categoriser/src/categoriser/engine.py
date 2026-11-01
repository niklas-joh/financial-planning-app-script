"""Turns transactions into categorised rows: rules first, then internal transfers.

Rows that neither covers are returned in `Result.uncovered_rows` for the model
(`model.py`), which labels only those it is confident about. Pure: no I/O.
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field

from .issues import Issue
from .merge import is_reviewed
from .rules import Rule, first_match, normalise, owner_slug
from .schema import CATEGORISED_COLUMNS, CategorisedRow, Method, transaction_context
from .taxonomy import Taxonomy
from .transfers import find_transfers, transfer_sub_category

TRANSFER_TYPE = TRANSFER_CATEGORY = "Internal Transfer"
SAMPLE_SIZE = 3
SAMPLE_MAX_CHARS = 300


@dataclass(frozen=True)
class RuleStats:
    hits: int          # rows the rule matches on its own (other rules and `active` ignored)
    wins: int          # rows the rule actually categorised (it was the first active match)
    lastMatched: str   # latest bookingDate among matched rows ('' if none)
    sample: str        # the most common counterparties among matched rows, see `sample_text`


@dataclass
class Result:
    rows: list[CategorisedRow] = field(default_factory=list)
    stats: dict[str, RuleStats] = field(default_factory=dict)  # every valid rule, active or not
    by_rule: int = 0
    by_transfer: int = 0
    uncovered: int = 0
    uncovered_rows: list[dict] = field(default_factory=list)  # for the model step


def sample_text(matched: list[dict]) -> str:
    """What a rule catches, e.g. `Acme Market (12); ACME MKT ONLINE (3); +2 more`.

    Rows are grouped by counterparty (description when the counterparty is blank),
    normalised as the rules match them; each group shows its most common spelling
    and row count, largest groups first.
    """
    groups: dict[str, Counter] = {}
    for txn in matched:
        raw = str(txn.get("counterparty", "") or "").strip() or str(txn.get("description", "") or "").strip()
        groups.setdefault(normalise(raw), Counter())[raw or "(blank)"] += 1
    ranked = sorted(groups.values(), key=lambda c: (-c.total(), c.most_common(1)[0][0].casefold()))
    parts = [f"{c.most_common(1)[0][0]} ({c.total()})" for c in ranked[:SAMPLE_SIZE]]
    if len(ranked) > SAMPLE_SIZE:
        parts.append(f"+{len(ranked) - SAMPLE_SIZE} more")
    text = "; ".join(parts)
    return text if len(text) <= SAMPLE_MAX_CHARS else text[: SAMPLE_MAX_CHARS - 1] + "…"


def rule_stats(rules: list[Rule], transactions: list[dict], wins: Counter | None = None) -> dict[str, RuleStats]:
    """How many transactions each rule matches on its own (ignoring other rules and `active`),
    so an inactive mined rule shows what it would do before it is approved. `wins` counts,
    per ruleId, the rows the rule actually categorised; `hits - wins` on an active rule is
    what a rule earlier in the matching order took instead."""
    wins = wins or Counter()
    stats = {}
    for rule in rules:
        matched = [t for t in transactions if rule.matches(t)]
        dates = [str(t.get("bookingDate", ""))[:10] for t in matched]
        stats[rule.ruleId] = RuleStats(len(matched), wins[rule.ruleId], max(dates, default=""),
                                       sample_text(matched))
    return stats


def categorise(transactions: list[dict], rules: list[Rule], taxonomy: Taxonomy,
               joint_owner: str, now: str) -> Result:
    """`rules` are the valid rules in matching order; only active ones categorise."""
    active = [r for r in rules if r.active]
    result = Result()
    wins: Counter = Counter()
    seen: set[str] = set()
    unique: list[dict] = []
    uncovered: list[dict] = []
    for txn in transactions:
        tid = str(txn.get("transactionId", "")).strip()
        if not tid or tid in seen:
            continue  # duplicate ids on `transactions` are bank-sync's to fix; first one wins
        seen.add(tid)
        unique.append(txn)
        rule = first_match(active, txn)
        if rule is None:
            uncovered.append(txn)
            continue
        owner = owner_slug(rule.owner) if rule.owner else owner_slug(txn.get("owner", ""))
        result.rows.append(CategorisedRow(tid, owner, rule.type, rule.category, rule.subCategory,
                                          Method.RULE, None, rule.ruleId, False, now, **transaction_context(txn)))
        result.by_rule += 1
        wins[rule.ruleId] += 1
    result.stats = rule_stats(rules, transactions, wins)

    # Pairs are looked for among all rows, so a leg pairs even when a rule covered the other leg;
    # only the legs no rule covered are labelled as transfers.
    sub_categories = {s for (t, c, s) in taxonomy.triples if t == TRANSFER_TYPE and c == TRANSFER_CATEGORY}
    transfers = find_transfers(unique) if sub_categories and uncovered else {}
    for txn in uncovered:
        tid = str(txn.get("transactionId", "")).strip()
        transfer = transfers.get(tid)
        if transfer is None:
            result.uncovered += 1
            result.uncovered_rows.append(txn)
            continue
        sub = transfer_sub_category(transfer, joint_owner, sub_categories)
        if (TRANSFER_TYPE, TRANSFER_CATEGORY, sub) not in taxonomy:
            result.uncovered += 1
            result.uncovered_rows.append(txn)
            continue
        result.rows.append(CategorisedRow(tid, owner_slug(txn.get("owner", "")), TRANSFER_TYPE,
                                          TRANSFER_CATEGORY, sub, Method.RULE, None, transfer.ruleId,
                                          False, now, **transaction_context(txn)))
        result.by_transfer += 1
    return result


REVIEW_DISAGREES = "rule_disagrees_with_review"
OUTPUT_FIELDS = ("owner", "type", "category", "subCategory")


def review_disagreements(categorised: list[list], rows: list[CategorisedRow], sheet: str) -> list[Issue]:
    """Reviewed rows are ground truth: report each one a rule would now label differently,
    so the rule can be fixed. `categorised` is the whole output sheet incl. header."""
    produced = {r.transactionId: r for r in rows}
    col = {name: i for i, name in enumerate(CATEGORISED_COLUMNS)}
    issues = []
    for row in categorised[1:]:
        row = list(row) + [""] * (len(CATEGORISED_COLUMNS) - len(row))
        tid = str(row[col["transactionId"]]).strip()
        if not tid or not is_reviewed(row[col["reviewed"]]) or tid not in produced:
            continue
        mine = produced[tid]
        reviewed = tuple(str(row[col[f]]).strip() for f in OUTPUT_FIELDS)
        ours = tuple(getattr(mine, f) for f in OUTPUT_FIELDS)
        if reviewed != ours:
            issues.append(Issue(
                type=REVIEW_DISAGREES, transactionId=tid,
                detail=f"Rule '{mine.ruleId}' gives {' / '.join(ours)}; the reviewed row says "
                       f"{' / '.join(reviewed)}. Fix the rule or add a more specific one.",
                source=f"categoriser.run: reviewed rows in '{sheet}'",
            ))
    return issues
