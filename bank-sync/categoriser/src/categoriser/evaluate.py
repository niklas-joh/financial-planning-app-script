"""Holdout accuracy of mined rules, transfer pairing and the model: `python -m categoriser.evaluate`.

Rules are mined, and the model trained, only from labelled examples dated before
the cutoff. The mined rules are all switched on and run over every transaction
together with the active rules already on the rules sheet; the model then labels
the rows they leave, at or above `--min-confidence` (default `MODEL_MIN_CONFIDENCE`).
Everything is scored on transactions with a legacy (or reviewed) label dated on or
after the cutoff. The report, with coverage and accuracy per confidence threshold
and example misses, goes to `<CATEGORISER_HOME>/data/eval/` and never into the
repo or the logs.
"""

from __future__ import annotations

import argparse
import logging
from collections import Counter
from dataclasses import dataclass, field, replace
from datetime import UTC, date, datetime

from .engine import categorise
from .legacy import LegacyRow
from .mine import LEGACY, REVIEWED, _example, labelled, mine_rules
from .model import REVIEWED_WEIGHT, Model, eligible
from .model import train as train_model
from .rules import Rule, account_label
from .taxonomy import Taxonomy, Triple
from .transfers import _day

DEFAULT_CUTOFF = "2026-03-01"
THRESHOLDS = (0.5, 0.6, 0.7, 0.8, 0.9, 0.95)
ALL = "all accounts"


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
    misses: list[tuple] = field(default_factory=list)          # (account, counterparty, description, predicted, actual, by)
    uncovered_labels: Counter = field(default_factory=Counter)  # actual labels of rows nothing covered
    min_confidence: float | None = None                        # None: no model was tried
    model_examples: int = 0
    # Per account (and ALL): (confidence, correct) of each eligible test row the rules leave.
    model_curve: dict[str, list[tuple[float, bool]]] = field(default_factory=dict)
    model_skipped_card: int = 0                                # test rows rules leave that are card rows


def weighted(items: list[tuple[dict, Triple, str]]) -> list[tuple[dict, Triple, float]]:
    """Model training examples from `mine.labelled` rows: reviewed rows weigh more."""
    return [(t, label, REVIEWED_WEIGHT if kind == REVIEWED else 1.0) for t, label, kind in items]


def evaluate(transactions: list[dict], legacy: list[LegacyRow], reviewed: dict[str, Triple], taxonomy: Taxonomy,
             manual_rules: list[Rule], joint_owner: str, cutoff: date,
             min_confidence: float | None = None) -> Evaluation:
    """With `min_confidence` set, a model trained before the cutoff labels what the rules leave."""
    items = [(t, label, kind) for t, label, kind in labelled(transactions, legacy, reviewed)
             if _day(t.get("bookingDate"))]
    train = [i for i in items if _day(i[0].get("bookingDate")) < cutoff]
    test = [(t, label) for t, label, kind in items if kind != LEGACY and _day(t.get("bookingDate")) >= cutoff]

    mined = [m.as_rule() for m in mine_rules(transactions, [_example(t, label, legacy=kind == LEGACY)
                                                            for t, label, kind in train], taxonomy)]
    rules = sorted([r for r in manual_rules if r.active] + mined, key=Rule.sort_key)
    result = categorise(transactions, rules, taxonomy, joint_owner, "eval")
    predicted = {r.transactionId: r for r in result.rows}

    ev = Evaluation(cutoff=cutoff, mined=len(mined), min_confidence=min_confidence)
    guesses: dict[str, tuple[Triple, float]] = {}
    if min_confidence is not None:
        left = [t for t, _label in test if str(t.get("transactionId", "")).strip() not in predicted]
        ev.model_skipped_card = sum(not eligible(t) for t in left)
        model = _train_or_none(weighted(train), taxonomy)
        if model is not None:
            ev.model_examples = model.meta["examples"]
            left = [t for t in left if eligible(t)]
            guesses = {str(t.get("transactionId", "")).strip(): g for t, g in zip(left, model.predict(left))}

    for txn, actual in sorted(test, key=lambda x: str(x[0].get("transactionId", ""))):
        tid = str(txn.get("transactionId", "")).strip()
        account = account_label(txn)
        row = predicted.get(tid)
        if row is not None:
            guess, kind, by = (row.type, row.category, row.subCategory), \
                ("transfer" if row.ruleId.startswith("transfer:") else "rule"), row.ruleId
        else:
            guess, kind, by = None, "none", ""
            if tid in guesses:
                triple, confidence = guesses[tid]
                for scope in (ALL, account):
                    ev.model_curve.setdefault(scope, []).append((confidence, triple == actual))
                if confidence >= min_confidence and triple in taxonomy:
                    guess, kind, by = triple, "model", f"model {confidence:.2f}"
        for score in (ev.overall, ev.by_method.setdefault(kind, Score()), ev.by_account.setdefault(account, Score())):
            score.add(guess, actual)
        if guess is None:
            ev.uncovered_labels[actual] += 1
        elif guess != actual:
            ev.misses.append((account, str(txn.get("counterparty", ""))[:40],
                              str(txn.get("description", ""))[:60], " / ".join(guess), " / ".join(actual), by))
    return ev


def _train_or_none(examples: list[tuple[dict, Triple, float]], taxonomy: Taxonomy) -> Model | None:
    try:
        return train_model(examples, taxonomy, "eval")
    except ValueError:
        return None  # too few labels before the cutoff


def curve(points: list[tuple[float, bool]]) -> dict[float, tuple[int, int, int]]:
    """threshold -> (rows labelled, of rows, exact) for the model alone."""
    return {th: (sum(c >= th for c, _ok in points), len(points), sum(c >= th and ok for c, ok in points))
            for th in THRESHOLDS}


def coverage_by_account(transactions: list[dict], rules: list[Rule], taxonomy: Taxonomy, joint_owner: str) -> dict:
    result = categorise(transactions, [replace(r, active=True) for r in rules], taxonomy, joint_owner, "eval")
    covered = {r.transactionId for r in result.rows}
    total, hit = Counter(), Counter()
    for t in transactions:
        total[account_label(t)] += 1
        hit[account_label(t)] += str(t.get("transactionId", "")).strip() in covered
    return {a: (hit[a], total[a]) for a in sorted(total)}


def render(ev: Evaluation, coverage: dict) -> str:
    model_note = ("" if ev.min_confidence is None else
                  f" The model ({ev.model_examples} examples) labels the rest at confidence ≥ {ev.min_confidence}.")
    lines = [f"# Holdout evaluation ({datetime.now(UTC):%Y-%m-%d %H:%M} UTC)", "",
             f"Rules mined from labelled rows before {ev.cutoff}: {ev.mined}.{model_note} Scored on labelled rows "
             f"from {ev.cutoff} on.", "", f"- **Overall:** {ev.overall.line()}"]
    lines += [f"- {k}: {s.line()}" for k, s in sorted(ev.by_method.items())]
    lines += ["", "## By account"] + [f"- {a}: {s.line()}" for a, s in sorted(ev.by_account.items())]
    if ev.min_confidence is not None:
        lines += render_curve(ev)
    lines += ["", "## Coverage if every rule mined from all data were switched on (all rows, no labels needed)"]
    lines += [f"- {a}: {h}/{n}" for a, (h, n) in coverage.items()]
    lines += ["", "## Labels of test rows nothing covered (top 15)"]
    lines += [f"- {n} × {' / '.join(lab)}" for lab, n in ev.uncovered_labels.most_common(15)]
    lines += ["", f"## Misses ({len(ev.misses)})", "", "| account | counterparty | description | predicted | actual | by |",
              "|---|---|---|---|---|---|"]
    lines += ["| " + " | ".join(str(c).replace("|", "/") for c in m) + " |" for m in ev.misses]
    return "\n".join(lines) + "\n"


def render_curve(ev: Evaluation) -> list[str]:
    """Model alone on the test rows rules leave: labelled share and accuracy per threshold."""
    scopes = [ALL] + sorted(k for k in ev.model_curve if k != ALL)
    lines = ["", "## Model on test rows the rules leave", "",
             f"Card rows are skipped by the model ({ev.model_skipped_card} such test rows). Each cell: rows labelled "
             "/ rows, and how many of those are exactly right.", "",
             "| confidence ≥ | " + " | ".join(scopes) + " |", "|---|" + "---|" * len(scopes)]
    curves = {scope: curve(ev.model_curve.get(scope, [])) for scope in scopes}
    for th in THRESHOLDS:
        cells = []
        for scope in scopes:
            n, total, exact = curves[scope][th]
            cells.append(f"{n}/{total}, {exact / n:.0%} exact" if n else f"0/{total}")
        lines.append(f"| {th} | " + " | ".join(cells) + " |")
    return lines


def load_inputs(settings) -> tuple:
    """(transactions, taxonomy, manual rules, legacy rows, reviewed labels) for the hand-run commands.
    The legacy sheet is opened read-only."""
    from .legacy import TRANSACTIONS_TAB, parse_legacy
    from .mine import reviewed_labels
    from .rules import owner_slug, parse_rules
    from .sheets import (ensure_categorised_sheet, ensure_rules_sheet, ensure_taxonomy_sheet, open_readonly,
                         open_spreadsheet, read_transactions, read_values)
    from .taxonomy import parse_taxonomy

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
    return transactions, taxonomy, manual, legacy, reviewed


def write_report(ev: Evaluation, transactions: list[dict], legacy: list[LegacyRow], reviewed: dict[str, Triple],
                 taxonomy: Taxonomy, manual: list[Rule], joint_owner: str):
    from .config import categoriser_home
    from .mine import build_examples

    # Rules mined from all labelled rows (not just before the cutoff, as in `evaluate`), for the coverage table.
    all_mined = [m.as_rule() for m in mine_rules(transactions, build_examples(transactions, legacy, reviewed),
                                                 taxonomy)]
    coverage = coverage_by_account(transactions, [r for r in manual if r.active] + all_mined, taxonomy, joint_owner)
    out = categoriser_home() / "data" / "eval"
    out.mkdir(parents=True, exist_ok=True)
    path = out / f"holdout-{datetime.now(UTC):%Y%m%d-%H%M%S}.md"
    path.write_text(render(ev, coverage))
    return path


def main() -> None:
    from .config import load_settings

    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--cutoff", default=DEFAULT_CUTOFF, help="first day of the test period (YYYY-MM-DD)")
    parser.add_argument("--min-confidence", type=float, help="model threshold (default MODEL_MIN_CONFIDENCE)")
    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

    settings = load_settings()
    transactions, taxonomy, manual, legacy, reviewed = load_inputs(settings)
    threshold = args.min_confidence if args.min_confidence is not None else settings.model_min_confidence
    ev = evaluate(transactions, legacy, reviewed, taxonomy, manual, settings.joint_owner,
                  date.fromisoformat(args.cutoff), threshold)
    path = write_report(ev, transactions, legacy, reviewed, taxonomy, manual, settings.joint_owner)
    s = ev.overall
    logging.getLogger("categoriser").info("test_rows=%d covered=%d exact=%d type_ok=%d report=%s",
                                          s.total, s.covered, s.exact, s.type_ok, path)


if __name__ == "__main__":
    main()
