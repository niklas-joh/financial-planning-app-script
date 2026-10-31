"""Entry point: `python -m categoriser.run` (what Docker and launchd run).

Reads transactions, the taxonomy and the rules, categorises with the active
rules and internal-transfer pairing, then lets the model (if one is trained)
label the confident rest. Writes `categorised`, refreshes the rules'
`hits`/`lastMatched` and logs problems in `categoriser_issues`. Logs counts
only, never transaction contents. A job sheet with renamed headers stops the run
before anything is written; a missing or broken model only skips the model step.
"""

from __future__ import annotations

import logging
from datetime import UTC, datetime

from .config import load_settings, models_dir
from .engine import REVIEW_DISAGREES, categorise, review_disagreements
from .issues import DUPLICATE_ID, duplicate_issues
from .mine import reviewed_labels
from .model import load as load_model
from .model import review_check
from .rules import INVALID_RULE, owner_slug, parse_rules
from .sheets import (
    ensure_categorised_sheet,
    ensure_issues_sheet,
    ensure_rules_sheet,
    ensure_taxonomy_sheet,
    open_spreadsheet,
    read_transactions,
    read_values,
    record_issues,
    upsert_categorised,
    write_rule_stats,
)
from .schema import Method
from .taxonomy import parse_taxonomy

log = logging.getLogger("categoriser")

CHECKED_TYPES = {DUPLICATE_ID, INVALID_RULE, REVIEW_DISAGREES}


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    settings = load_settings()
    ss = open_spreadsheet(settings)
    ws = ensure_categorised_sheet(ss, settings)
    issues_ws = ensure_issues_sheet(ss, settings)
    taxonomy_ws = ensure_taxonomy_sheet(ss, settings)
    rules_ws = ensure_rules_sheet(ss, settings)
    transactions = read_transactions(ss, settings)
    now = datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")

    owners = {owner_slug(t.get("owner", "")) for t in transactions} - {""}
    taxonomy = parse_taxonomy(read_values(taxonomy_ws))
    parsed = parse_rules(read_values(rules_ws), taxonomy, owners, settings.rules_sheet)
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
    # Read for the reviewed rows only (the upsert never changes those); the upsert re-reads the sheet.
    before_upsert = read_values(ws)
    reviewed_checked, reviewed_wrong = review_check(model_rows, reviewed_labels(before_upsert))
    plan = upsert_categorised(ws, settings, result.rows + model_rows, clear_stale_at=now,
                              clear_methods=clear_methods)
    stats_written = write_rule_stats(rules_ws, settings, result.stats)

    found = (parsed.issues + review_disagreements(before_upsert, result.rows, settings.output_sheet)
             + duplicate_issues(plan.duplicates, settings.output_sheet))
    issues = record_issues(issues_ws, settings, found, CHECKED_TYPES, now)
    log.info(
        "read=%d taxonomy=%d rules_active=%d rules_invalid=%d by_rule=%d by_transfer=%d uncovered=%d "
        "model=%s by_model=%d model_reviewed_checked=%d model_reviewed_wrong=%d "
        "updated=%d appended=%d cleared=%d skipped_reviewed=%d skipped_moved=%d skipped_duplicate=%d "
        "rule_stats_written=%d issues_opened=%d issues_reopened=%d issues_resolved=%d",
        len(transactions), len(taxonomy), len(parsed.active), len(parsed.issues),
        result.by_rule, result.by_transfer, result.uncovered,
        model_status, len(model_rows), reviewed_checked, reviewed_wrong,
        len(plan.updates) - len(plan.cleared), len(plan.appends), len(plan.cleared),
        len(plan.skipped_reviewed), len(plan.skipped_moved), len(plan.duplicates), stats_written,
        issues.opened, issues.reopened, issues.resolved,
    )


if __name__ == "__main__":
    main()
