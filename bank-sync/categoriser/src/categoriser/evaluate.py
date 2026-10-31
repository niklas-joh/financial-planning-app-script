"""Holdout accuracy of mined rules + transfer pairing: `python -m categoriser.evaluate [--cutoff YYYY-MM-DD]`.

Rules are mined only from labelled examples dated before the cutoff, all switched
on, and run over every transaction together with the active rules already on the
rules sheet. They are scored on transactions joined to a legacy (or reviewed)
label dated on or after the cutoff. The report, with example misses, goes to
`<CATEGORISER_HOME>/data/eval/` and never into the repo or the logs.
"""

from __future__ import annotations

import argparse
import logging
from collections import Counter
from dataclasses import dataclass, field, replace
from datetime import UTC, date, datetime

from .engine import categorise
from .legacy import LegacyRow, join_legacy
from .mine import Example, _example, mine_rules
from .rules import Rule, account_label
from .taxonomy import Taxonomy, Triple
from .transfers import _day

DEFAULT_CUTOFF = "2026-03-01"


@dataclass
class Score:
    total: int = 0
    covered: int = 0
    exact: int = 0
    type_ok: int = 0

    def add(self, predicted: Triple | None, actual: Triple) -> None:
        self.total += 1
        if predicted is None:
            return
        self.covered += 1
        self.exact += predicted == actual
        self.type_ok += predicted[0] == actual[0]

    def line(self) -> str:
        pct = lambda a, b: f"{a / b:.0%}" if b else "-"
        return (f"{self.total} rows, covered {self.covered} ({pct(self.covered, self.total)}), "
                f"exact {pct(self.exact, self.covered)}, type {pct(self.type_ok, self.covered)} of covered")


@dataclass
class Evaluation:
    cutoff: date
    mined: int
    overall: Score = field(default_factory=Score)
    by_method: dict[str, Score] = field(default_factory=dict)
    by_account: dict[str, Score] = field(default_factory=dict)
    misses: list[tuple] = field(default_factory=list)          # (account, counterparty, description, predicted, actual, ruleId)
    uncovered_labels: Counter = field(default_factory=Counter)  # actual labels of rows nothing covered


def evaluate(transactions: list[dict], legacy: list[LegacyRow], reviewed: dict[str, Triple], taxonomy: Taxonomy,
             manual_rules: list[Rule], joint_owner: str, cutoff: date) -> Evaluation:
    joined = join_legacy(legacy, transactions)
    used = {id(r) for r in joined.values()}
    truth: dict[str, tuple[date, Triple]] = {tid: (r.day, r.label) for tid, r in joined.items()}
    for t in transactions:
        tid = str(t.get("transactionId", "")).strip()
        day = _day(t.get("bookingDate"))
        if tid in reviewed and day:
            truth[tid] = (day, reviewed[tid])

    by_id = {str(t.get("transactionId", "")).strip(): t for t in transactions}
    train: list[Example] = [_example(by_id[tid], label) for tid, (day, label) in truth.items() if day < cutoff]
    train += [_example(r.as_transaction(), r.label, legacy=True)
              for r in legacy if id(r) not in used and r.day < cutoff]

    mined = [m.as_rule() for m in mine_rules(transactions, train, taxonomy)]
    rules = sorted([r for r in manual_rules if r.active] + mined, key=Rule.sort_key)
    result = categorise(transactions, rules, taxonomy, joint_owner, "eval")
    predicted = {r.transactionId: r for r in result.rows}

    ev = Evaluation(cutoff=cutoff, mined=len(mined))
    for tid, (day, actual) in sorted(truth.items()):
        if day < cutoff:
            continue
        row = predicted.get(tid)
        guess = (row.type, row.category, row.subCategory) if row else None
        txn = by_id[tid]
        kind = "none" if row is None else ("transfer" if row.ruleId.startswith("transfer:") else "rule")
        for score in (ev.overall, ev.by_method.setdefault(kind, Score()),
                      ev.by_account.setdefault(account_label(txn), Score())):
            score.add(guess, actual)
        if row is None:
            ev.uncovered_labels[actual] += 1
        elif guess != actual:
            ev.misses.append((account_label(txn), str(txn.get("counterparty", ""))[:40],
                              str(txn.get("description", ""))[:60], " / ".join(guess), " / ".join(actual), row.ruleId))
    return ev


def coverage_by_account(transactions: list[dict], rules: list[Rule], taxonomy: Taxonomy, joint_owner: str) -> dict:
    result = categorise(transactions, [replace(r, active=True) for r in rules], taxonomy, joint_owner, "eval")
    covered = {r.transactionId for r in result.rows}
    total, hit = Counter(), Counter()
    for t in transactions:
        total[account_label(t)] += 1
        hit[account_label(t)] += str(t.get("transactionId", "")).strip() in covered
    return {a: (hit[a], total[a]) for a in sorted(total)}


def render(ev: Evaluation, coverage: dict) -> str:
    lines = [f"# Holdout evaluation ({datetime.now(UTC):%Y-%m-%d %H:%M} UTC)", "",
             f"Rules mined from labelled rows before {ev.cutoff}: {ev.mined}. Scored on labelled rows from "
             f"{ev.cutoff} on.", "", f"- **Overall:** {ev.overall.line()}"]
    lines += [f"- {k}: {s.line()}" for k, s in sorted(ev.by_method.items())]
    lines += ["", "## By account"] + [f"- {a}: {s.line()}" for a, s in sorted(ev.by_account.items())]
    lines += ["", "## Coverage if every rule mined from all data were switched on (all rows, no labels needed)"]
    lines += [f"- {a}: {h}/{n}" for a, (h, n) in coverage.items()]
    lines += ["", "## Labels of test rows nothing covered (top 15)"]
    lines += [f"- {n} × {' / '.join(lab)}" for lab, n in ev.uncovered_labels.most_common(15)]
    lines += ["", f"## Misses ({len(ev.misses)})", "", "| account | counterparty | description | predicted | actual | rule |",
              "|---|---|---|---|---|---|"]
    lines += ["| " + " | ".join(str(c).replace("|", "/") for c in m) + " |" for m in ev.misses]
    return "\n".join(lines) + "\n"


def main() -> None:
    from .config import categoriser_home, load_settings
    from .legacy import TRANSACTIONS_TAB, parse_legacy
    from .mine import build_examples, reviewed_labels
    from .rules import owner_slug, parse_rules
    from .sheets import (ensure_categorised_sheet, ensure_rules_sheet, ensure_taxonomy_sheet, open_readonly,
                         open_spreadsheet, read_transactions, read_values)
    from .taxonomy import parse_taxonomy

    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--cutoff", default=DEFAULT_CUTOFF, help="first day of the test period (YYYY-MM-DD)")
    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

    settings = load_settings()
    if not settings.legacy_spreadsheet_id:
        raise SystemExit("LEGACY_SPREADSHEET_ID is not set in .env.")
    ss = open_spreadsheet(settings)
    transactions = read_transactions(ss, settings)
    taxonomy = parse_taxonomy(read_values(ensure_taxonomy_sheet(ss, settings)))
    owners = {owner_slug(t.get("owner", "")) for t in transactions} - {""}
    manual = [r for r in parse_rules(read_values(ensure_rules_sheet(ss, settings)), taxonomy, owners,
                                     settings.rules_sheet).rules if r.source != "mined"]
    legacy = parse_legacy(read_values(open_readonly(settings, settings.legacy_spreadsheet_id)
                                      .worksheet(TRANSACTIONS_TAB)))
    reviewed = reviewed_labels(read_values(ensure_categorised_sheet(ss, settings)))

    ev = evaluate(transactions, legacy, reviewed, taxonomy, manual, settings.joint_owner,
                  date.fromisoformat(args.cutoff))
    all_mined = [m.as_rule() for m in mine_rules(transactions, build_examples(transactions, legacy, reviewed),
                                                 taxonomy)]
    coverage = coverage_by_account(transactions, [r for r in manual if r.active] + all_mined, taxonomy,
                                   settings.joint_owner)
    out = categoriser_home() / "data" / "eval"
    out.mkdir(parents=True, exist_ok=True)
    path = out / f"holdout-{datetime.now(UTC):%Y%m%d-%H%M%S}.md"
    path.write_text(render(ev, coverage))
    s = ev.overall
    logging.getLogger("categoriser").info("test_rows=%d covered=%d exact=%d type_ok=%d report=%s",
                                          s.total, s.covered, s.exact, s.type_ok, path)


if __name__ == "__main__":
    main()
