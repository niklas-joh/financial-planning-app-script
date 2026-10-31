"""Train the model: `python -m categoriser.train [--cutoff YYYY-MM-DD]`. Run by hand, never by the schedule.

1. Holdout: the same check as `categoriser.evaluate` (rules and a model built only
   from rows before the cutoff, scored on labelled rows after it). The report goes
   to `<CATEGORISER_HOME>/data/eval/`.
2. Trains on every labelled row (legacy rows, Bank Sync rows joined to them, rows
   you reviewed) and saves the model with a small JSON description to
   `<CATEGORISER_HOME>/models/`. The next daily run uses it.

Run it with the same image as the schedule (`docker compose run --rm categoriser
python -m categoriser.train`): a model saved by another scikit-learn version is
skipped by the daily run as incompatible.
"""

from __future__ import annotations

import argparse
import logging
from collections import Counter
from datetime import UTC, date, datetime

from .config import load_settings, models_dir
from .evaluate import ALL, DEFAULT_CUTOFF, curve, evaluate, load_inputs, weighted, write_report
from .mine import labelled
from .model import save, train


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--cutoff", default=DEFAULT_CUTOFF, help="first day of the holdout period (YYYY-MM-DD)")
    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    log = logging.getLogger("categoriser")

    settings = load_settings()
    transactions, taxonomy, manual, legacy, reviewed = load_inputs(settings)
    threshold = settings.model_min_confidence
    ev = evaluate(transactions, legacy, reviewed, taxonomy, manual, settings.joint_owner,
                  date.fromisoformat(args.cutoff), threshold)
    report = write_report(ev, transactions, legacy, reviewed, taxonomy, manual, settings.joint_owner)

    items = labelled(transactions, legacy, reviewed)
    try:
        model = train(weighted(items), taxonomy, datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ"))
    except ValueError as error:
        raise SystemExit(f"No model saved: {error}") from None
    model.meta["examplesByKind"] = dict(Counter(kind for _t, _l, kind in items))
    model.meta["holdout"] = {
        "cutoff": args.cutoff, "minConfidence": threshold,
        "overall": vars(ev.overall),
        "model": {str(th): dict(zip(("labelled", "rows", "exact"), v))
                  for th, v in curve(ev.model_curve.get(ALL, [])).items()},
    }
    path = save(model, models_dir())
    m = ev.by_method.get("model")
    log.info("examples=%d labels=%d holdout_rows=%d holdout_by_model=%d holdout_model_exact=%d model=%s report=%s",
             model.meta["examples"], model.meta["labels"], ev.overall.total, m.covered if m else 0,
             m.exact if m else 0, path, report)


if __name__ == "__main__":
    main()
