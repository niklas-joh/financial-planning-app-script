"""Entry point: `python -m categoriser.run` (what Docker and launchd run).

Applies the answers ticked in the review inbox (reviewed `manual` rows, and rules
for "all like this"), reads transactions, the taxonomy and the rules, categorises
with the active rules and internal-transfer pairing, then lets the model (if one
is trained) label the confident rest. Writes `categorised`, refreshes the rules'
`hits`/`lastMatched`, fills the review inbox again and logs problems in
`categoriser_issues`. Logs counts only, never transaction contents. A job sheet
with renamed headers stops the run before anything is written; a missing or
broken model only skips the model step.
"""

from __future__ import annotations

import json
import logging
import random
from datetime import UTC, datetime
from pathlib import Path

from .config import data_dir, load_settings, models_dir
from .engine import REVIEW_DISAGREES, categorise, review_disagreements
from .issues import DUPLICATE_ID, duplicate_issues
from .merge import is_reviewed
from .mine import reviewed_labels
from .model import load as load_model
from .model import review_check
from .review import INVALID_REVIEW, REVIEW_CONFLICT, Guess, answer_text, apply_answers, plan_inbox, ticked
from .rules import INVALID_RULE, owner_slug, parse_rules
from .sheets import (
    append_rows,
    ensure_categorised_sheet,
    ensure_issues_sheet,
    ensure_review_sheet,
    ensure_rules_sheet,
    ensure_taxonomy_sheet,
    open_spreadsheet,
    read_transactions,
    read_values,
    record_issues,
    upsert_categorised,
    write_review,
    write_rule_stats,
)
from .schema import CATEGORISED_COLUMNS, Method
from .taxonomy import parse_taxonomy

log = logging.getLogger("categoriser")

CHECKED_TYPES = {DUPLICATE_ID, INVALID_RULE, REVIEW_DISAGREES, INVALID_REVIEW, REVIEW_CONFLICT}
HISTORY_FILE = "review-history.jsonl"


def _append_history(path: Path, entries: list[dict]) -> None:
    """One JSON object per answer applied, for measuring the model later (runtime folder only)."""
    if not entries:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as f:
        for entry in entries:
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")


def _after_write(categorised: list[list]) -> tuple[set[str], set[str]]:
    """(ids with a full type / category / subCategory, ids of unreviewed `model` rows) on
    `categorised` as written, so spot checks never pick a row the owner already answered."""
    col = {n: i for i, n in enumerate(CATEGORISED_COLUMNS)}
    labelled, by_model = set(), set()
    for row in categorised[1:]:
        row = list(row) + [""] * (len(CATEGORISED_COLUMNS) - len(row))
        tid = str(row[col["transactionId"]]).strip()
        if tid and all(str(row[col[n]]).strip() for n in ("type", "category", "subCategory")):
            labelled.add(tid)
            if str(row[col["method"]]).strip() == Method.MODEL and not is_reviewed(row[col["reviewed"]]):
                by_model.add(tid)
    return labelled, by_model


def _safe_guess(model) -> Guess | None:
    """The model's guesses for the inbox; a failure only leaves the answers blank."""
    if model is None:
        return None

    def guess(txns: list[dict]) -> list:
        try:
            return model.predict(txns)
        except Exception:  # noqa: BLE001 - pre-filling the inbox must never stop the run
            log.exception("model guess for the review inbox failed; leaving answers blank")
            return [None] * len(txns)
    return guess


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    settings = load_settings()
    ss = open_spreadsheet(settings)
    ws = ensure_categorised_sheet(ss, settings)
    issues_ws = ensure_issues_sheet(ss, settings)
    taxonomy_ws = ensure_taxonomy_sheet(ss, settings)
    rules_ws = ensure_rules_sheet(ss, settings)
    review_ws = ensure_review_sheet(ss, settings)
    transactions = read_transactions(ss, settings)
    now = datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
    today = now[:10]

    owners = {owner_slug(t.get("owner", "")) for t in transactions} - {""}
    taxonomy = parse_taxonomy(read_values(taxonomy_ws))
    # Read for the reviewed rows only (the upsert never changes those); the upsert re-reads the sheet.
    before_upsert = read_values(ws)
    review_values = read_values(review_ws)
    rules_values = read_values(rules_ws)

    # Answers first, so a rule made with "all like this" already applies in this run.
    by_id: dict[str, dict] = {}
    for t in transactions:
        by_id.setdefault(str(t.get("transactionId", "")).strip(), t)
    rule_ids = {str(r[0]).strip() for r in rules_values[1:] if r and str(r[0]).strip()}
    answers = apply_answers(review_values, by_id, taxonomy, owners, before_upsert, rule_ids, today, now,
                            settings.review_sheet)
    if answers.new_rules:
        append_rows(rules_ws, settings, answers.new_rules)
        rules_values = read_values(rules_ws)

    parsed = parse_rules(rules_values, taxonomy, owners, settings.rules_sheet)
    result = categorise(transactions, parsed.rules, taxonomy, settings.joint_owner, now)

    model, model_status = load_model(models_dir())
    model_rows = []
    if model is not None:
        try:
            model_rows = model.rows(result.uncovered_rows, taxonomy, settings.model_min_confidence, now)
        except Exception:  # noqa: BLE001 - the model step must never stop the run
            log.exception("model step failed; continuing without it")
            model, model_status = None, "failed"
    # Without a working model, keep the model rows already on the sheet instead of clearing them.
    clear_methods = frozenset({Method.RULE, Method.MODEL} if model is not None else {Method.RULE})
    reviewed_checked, reviewed_wrong = review_check(model_rows, reviewed_labels(before_upsert))
    manual_ids = {r.transactionId for r in answers.manual_rows}
    produced = [r for r in result.rows + model_rows if r.transactionId not in manual_ids]
    plan = upsert_categorised(ws, settings, produced + answers.manual_rows, clear_stale_at=now,
                              clear_methods=clear_methods)
    stats_written = write_rule_stats(rules_ws, settings, result.stats)
    _append_history(data_dir() / HISTORY_FILE, answers.history)

    labelled, by_model = _after_write(read_values(ws))
    inbox = plan_inbox(review_values, answers, transactions, labelled, by_model, taxonomy, _safe_guess(model),
                       settings.review_size, settings.review_spot_checks, random.Random(today))
    review_written = write_review(review_ws, settings, inbox, ticked(review_values),
                                  [answer_text(t) for t in sorted(taxonomy.triples)])

    found = (parsed.issues + review_disagreements(before_upsert, result.rows, settings.output_sheet)
             + duplicate_issues(plan.duplicates, settings.output_sheet) + answers.issues)
    issues = record_issues(issues_ws, settings, found, CHECKED_TYPES, now)
    log.info(
        "read=%d taxonomy=%d rules_active=%d rules_invalid=%d by_rule=%d by_transfer=%d uncovered=%d "
        "model=%s by_model=%d model_reviewed_checked=%d model_reviewed_wrong=%d "
        "updated=%d appended=%d cleared=%d skipped_reviewed=%d skipped_moved=%d skipped_duplicate=%d "
        "rule_stats_written=%d review_answers=%d review_invalid=%d review_rules_added=%d review_changed=%d "
        "review_spot_checks_right=%d review_spot_checks_wrong=%d review_size=%d review=%s "
        "issues_opened=%d issues_reopened=%d issues_resolved=%d",
        len(transactions), len(taxonomy), len(parsed.active), len(parsed.issues),
        result.by_rule, result.by_transfer, result.uncovered,
        model_status, len(model_rows), reviewed_checked, reviewed_wrong,
        len(plan.updates) - len(plan.cleared), len(plan.appends), len(plan.cleared),
        len(plan.skipped_reviewed), len(plan.skipped_moved), len(plan.duplicates), stats_written,
        len(answers.manual_rows), len(answers.keep), len(answers.new_rules), answers.changed,
        answers.spot_right, answers.spot_wrong, len(inbox), "written" if review_written else "skipped_busy",
        issues.opened, issues.reopened, issues.resolved,
    )


if __name__ == "__main__":
    main()
