"""Mine rule suggestions from labelled history: `python -m categoriser.mine [--dry-run]`.

Run by hand, not by the daily job. Labelled examples come from
- legacy rows joined to a Bank Sync transaction (that transaction's own text),
- legacy rows that join to nothing (their own text; other banks' formats), and
- rows marked reviewed in `categorised` (ground truth).

For each kind of transaction on `transactions` the miner tries, in order, word
sequences of the counterparty (1 to 3 words, digits dropped, longest first), then
of the description, then the provider category, and keeps the first one whose
labelled examples agree enough (`MIN_EXAMPLES`, `MIN_AGREEMENT`). New rules are
appended with `active` = FALSE and `source` = mined, for the owner to approve.
Existing rules are never changed, and a transaction an existing rule's text
already names gets no new suggestion.
"""

from __future__ import annotations

import argparse
import hashlib
import logging
from collections import Counter
from dataclasses import dataclass
from datetime import UTC, datetime

from .legacy import TRANSACTIONS_TAB, LegacyRow, join_legacy, parse_legacy
from .merge import is_reviewed
from .rules import DEFAULT_PRIORITY, RCOL, RULE_COLUMNS, Rule, amount_eur, normalise
from .schema import CATEGORISED_COLUMNS
from .taxonomy import Taxonomy, Triple

log = logging.getLogger("categoriser")

MIN_EXAMPLES = 3
MIN_AGREEMENT = 0.9
MIN_NEEDLE_LENGTH = 4
MAX_NEEDLE_WORDS = 3


TEXT_FIELDS = ("counterpartyContains", "descriptionContains")  # tried in this order, then providerCategory


@dataclass(frozen=True)
class Example:
    counterparty: str       # normalised
    description: str        # normalised
    provider_category: str  # normalised provider category; '' for legacy-only rows
    direction: str          # in / out
    label: Triple


@dataclass(frozen=True)
class MinedRule:
    field: str         # counterpartyContains, descriptionContains or providerCategory
    needle: str
    direction: str     # '' when the agreeing examples go both ways
    label: Triple
    examples: int
    agreement: float

    @property
    def rule_id(self) -> str:
        return "m-" + hashlib.sha1(f"{self.field}|{self.needle}|{self.direction}".encode()).hexdigest()[:8]

    def as_rule(self, active: bool = True) -> Rule:
        return Rule(ruleId=self.rule_id, active=active, priority=DEFAULT_PRIORITY, direction=self.direction,
                    type=self.label[0], category=self.label[1], subCategory=self.label[2], source="mined",
                    **{self.field: self.needle})

    def to_values(self, today: str) -> list:
        values = [""] * len(RULE_COLUMNS)
        values[RCOL["ruleId"]] = self.rule_id
        values[RCOL["active"]] = False
        values[RCOL["priority"]] = int(DEFAULT_PRIORITY)
        values[RCOL["direction"]] = self.direction
        values[RCOL[self.field]] = self.needle
        values[RCOL["type"]], values[RCOL["category"]], values[RCOL["subCategory"]] = self.label
        values[RCOL["notes"]] = (f"mined {today}: {self.examples} labelled rows, "
                                 f"{round(self.agreement * 100)}% agree")
        values[RCOL["source"]] = "mined"
        return values


def _example(txn: dict, label: Triple, legacy: bool = False) -> Example:
    counterparty, description = normalise(txn.get("counterparty", "")), normalise(txn.get("description", ""))
    if legacy:
        # Legacy columns map only roughly (the operation text often holds the name): use both.
        counterparty = description = f"{counterparty} {description}".strip()
    return Example(counterparty, description, normalise(txn.get("category", "")),
                   "in" if amount_eur(txn) > 0 else "out", label)


def reviewed_labels(categorised: list[list]) -> dict[str, Triple]:
    """transactionId -> (type, category, subCategory) of rows marked reviewed."""
    col = {n: i for i, n in enumerate(CATEGORISED_COLUMNS)}
    out = {}
    for row in categorised[1:]:
        row = list(row) + [""] * (len(CATEGORISED_COLUMNS) - len(row))
        label = tuple(str(row[col[n]]).strip() for n in ("type", "category", "subCategory"))
        if is_reviewed(row[col["reviewed"]]) and all(label) and str(row[col["transactionId"]]).strip():
            out[str(row[col["transactionId"]]).strip()] = label
    return out


def build_examples(transactions: list[dict], legacy: list[LegacyRow], reviewed: dict[str, Triple]) -> list[Example]:
    """Reviewed rows win over a joined legacy label for the same transaction."""
    joined = join_legacy(legacy, transactions)
    used = {id(r) for r in joined.values()}
    examples = []
    for t in transactions:
        tid = str(t.get("transactionId", "")).strip()
        if tid in reviewed:
            examples.append(_example(t, reviewed[tid]))
        elif tid in joined:
            examples.append(_example(t, joined[tid].label))
    examples += [_example(r.as_transaction(), r.label, legacy=True) for r in legacy if id(r) not in used]
    return examples


def _needles(text: object) -> list[str]:
    """Word sequences (1-3 words, no digits), longest first."""
    words = [w for w in normalise(text).split() if not any(c.isdigit() for c in w)]
    out = set()
    for size in range(min(MAX_NEEDLE_WORDS, len(words)), 0, -1):
        for i in range(len(words) - size + 1):
            needle = " ".join(words[i:i + size])
            if len(needle) >= MIN_NEEDLE_LENGTH:
                out.add(needle)
    return sorted(out, key=lambda n: (-len(n.split()), -len(n), n))


def _contains(text: str, needle: str) -> bool:
    """`text` and `needle` are already normalised (same test as `rules.contains_words`)."""
    return f" {needle} " in f" {text} "


def _judge(field: str, needle: str, examples: list[Example], taxonomy: Taxonomy) -> MinedRule | None:
    if field == "providerCategory":
        hits = [e for e in examples if e.provider_category == needle]
    elif field == "counterpartyContains":
        hits = [e for e in examples if _contains(e.counterparty, needle)]
    else:
        hits = [e for e in examples if _contains(e.description, needle)]
    if len(hits) < MIN_EXAMPLES:
        return None
    label, n = Counter(e.label for e in hits).most_common(1)[0]
    if n / len(hits) < MIN_AGREEMENT or label not in taxonomy:
        return None
    directions = {e.direction for e in hits if e.label == label}
    return MinedRule(field, needle, directions.pop() if len(directions) == 1 else "", label, len(hits), n / len(hits))


def mine_rules(transactions: list[dict], examples: list[Example], taxonomy: Taxonomy,
               existing: set[tuple[str, str]] = frozenset()) -> list[MinedRule]:
    """At most one suggestion per distinct (counterparty, description, provider category)
    on `transactions`: the first agreeing needle, counterparty words first (longest
    first), then description words, then the provider category. Pure.

    A suggestion must match the transaction it came from (an out-going rule mined
    for an incoming row is skipped). `existing` holds (field, normalised value) of
    rules already on the sheet; a transaction one of them would cover gets no
    suggestion.
    """
    shapes: dict[tuple[str, str, str], dict] = {}
    for t in transactions:
        shape = (str(t.get("counterparty", "")), str(t.get("description", "")), str(t.get("category", "")))
        shapes.setdefault(shape, t)
    chosen: dict[tuple[str, str], MinedRule] = {}
    cache: dict[tuple[str, str], MinedRule | None] = {}
    for (counterparty, description, provider_category), sample in sorted(shapes.items(), key=lambda kv: kv[0]):
        candidates = [("counterpartyContains", n) for n in _needles(counterparty)]
        candidates += [("descriptionContains", n) for n in _needles(description)]
        if normalise(provider_category):
            candidates.append(("providerCategory", normalise(provider_category)))
        if any(c in existing for c in candidates):
            continue
        for key in candidates:
            if key not in cache:
                cache[key] = _judge(*key, examples, taxonomy)
            # The suggestion must cover the transaction it came from (its direction may not).
            if cache[key] is not None and cache[key].as_rule().matches(sample):
                chosen.setdefault(key, cache[key])
                break
    return sorted(chosen.values(), key=lambda r: (r.field, r.needle))


def existing_signatures(rules_values: list[list]) -> set[tuple[str, str]]:
    """(field, normalised value) of every rule's text/provider-category match columns."""
    out = set()
    for row in rules_values[1:]:
        for field in (*TEXT_FIELDS, "providerCategory"):
            i = RCOL[field]
            if len(row) > i and normalise(row[i]):
                out.add((field, normalise(row[i])))
    return out


def main() -> None:
    from .config import load_settings
    from .sheets import (append_rows, ensure_categorised_sheet, ensure_rules_sheet, ensure_taxonomy_sheet,
                         open_readonly, open_spreadsheet, read_transactions, read_values)
    from .taxonomy import parse_taxonomy

    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--dry-run", action="store_true", help="count suggestions, write nothing")
    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

    settings = load_settings()
    if not settings.legacy_spreadsheet_id:
        raise SystemExit("LEGACY_SPREADSHEET_ID is not set in .env.")
    ss = open_spreadsheet(settings)
    transactions = read_transactions(ss, settings)
    taxonomy = parse_taxonomy(read_values(ensure_taxonomy_sheet(ss, settings)))
    rules_ws = ensure_rules_sheet(ss, settings)
    legacy = parse_legacy(read_values(open_readonly(settings, settings.legacy_spreadsheet_id)
                                      .worksheet(TRANSACTIONS_TAB)))
    reviewed = reviewed_labels(read_values(ensure_categorised_sheet(ss, settings)))
    examples = build_examples(transactions, legacy, reviewed)
    rules_values = read_values(rules_ws)
    known_ids = {str(r[0]).strip() for r in rules_values[1:] if r}
    mined = [m for m in mine_rules(transactions, examples, taxonomy, existing_signatures(rules_values))
             if m.rule_id not in known_ids]
    today = datetime.now(UTC).date().isoformat()
    if not args.dry_run and mined:
        append_rows(rules_ws, settings, [m.to_values(today) for m in mined])
    log.info("examples=%d (reviewed=%d) suggestions=%d written=%s",
             len(examples), len(reviewed), len(mined), not args.dry_run)


if __name__ == "__main__":
    main()
