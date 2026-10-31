"""Text classifier for rows no rule or transfer pairing covers (phase 2).

Logistic regression over the words and letter n-grams of counterparty, description
and reference, plus direction and the size of the amount. It predicts a
(type, category, subCategory) triple with a probability; the daily run writes it
only at or above `MODEL_MIN_CONFIDENCE`, so a weak guess leaves the row blank.

Card rows (`classification` = card) are skipped: the hand-labelled history has
too few comparable card rows to be trusted (see the README).

Trained by hand with `python -m categoriser.train`; the model and a small JSON
file describing it live in `<CATEGORISER_HOME>/models/`, never in the repo.
"""

from __future__ import annotations

import json
import math
import re
from collections import Counter
from dataclasses import dataclass
from pathlib import Path

from .rules import amount_eur, owner_slug
from .schema import CategorisedRow, Method
from .taxonomy import Taxonomy, Triple

MODEL_FILE = "categoriser-model.joblib"
META_FILE = "categoriser-model.json"
FORMAT = 1
REVIEWED_WEIGHT = 3.0  # rows you ticked reviewed count three times
MIN_LABEL_EXAMPLES = 2  # labels with fewer examples are not learnt
NOISE_WORDS = frozenset({"withdraw", "deposit"})  # Trading 212 puts these in every description
_SEP = "\t"  # between type, category and subCategory in class names ("Wants/Pleasure" holds a slash)


def model_text(txn: dict) -> str:
    """Words of counterparty, description and reference, without numbers (dates, card
    numbers, references) and one-letter words."""
    text = re.sub(r"[\W_]+", " ", " ".join(str(txn.get(k, "") or "") for k in
                                            ("counterparty", "description", "reference")).casefold())
    return " ".join(w for w in text.split()
                    if len(w) > 1 and w not in NOISE_WORDS and not any(c.isdigit() for c in w))


def features(txn: dict) -> dict:
    amount = amount_eur(txn)
    return {"text": model_text(txn), "direction": "in" if amount > 0 else "out",
            "size": str(int(math.log2(abs(amount) + 1)))}


def eligible(txn: dict) -> bool:
    """Rows the model may label: not card rows."""
    return str(txn.get("classification", "")).strip().lower() != "card"


def _frame(txns: list[dict]):
    import pandas as pd
    return pd.DataFrame([features(t) for t in txns], columns=["text", "direction", "size"])


def _pipeline():
    from sklearn.compose import ColumnTransformer
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.linear_model import LogisticRegression
    from sklearn.pipeline import Pipeline
    from sklearn.preprocessing import OneHotEncoder

    columns = ColumnTransformer([
        ("words", TfidfVectorizer(ngram_range=(1, 2), sublinear_tf=True), "text"),
        ("chars", TfidfVectorizer(analyzer="char_wb", ngram_range=(3, 5), sublinear_tf=True, min_df=2), "text"),
        ("other", OneHotEncoder(handle_unknown="ignore"), ["direction", "size"]),
    ])
    return Pipeline([("features", columns), ("classifier", LogisticRegression(C=20, max_iter=3000))])


@dataclass
class Model:
    pipeline: object
    meta: dict

    def predict(self, txns: list[dict]) -> list[tuple[Triple, float]]:
        """(triple, probability) per transaction, in order."""
        if not txns:
            return []
        proba = self.pipeline.predict_proba(_frame(txns))
        classes = self.pipeline.classes_
        out = []
        for row in proba:
            best = int(row.argmax())
            out.append((tuple(str(classes[best]).split(_SEP)), float(row[best])))
        return out

    def rows(self, txns: list[dict], taxonomy: Taxonomy, min_confidence: float, now: str) -> list[CategorisedRow]:
        """`method` = model rows for the eligible `txns` predicted at or above `min_confidence`
        with a triple still in the taxonomy. Owner is the account's owner."""
        txns = [t for t in txns if eligible(t)]
        out = []
        for txn, (triple, confidence) in zip(txns, self.predict(txns)):
            if confidence >= min_confidence and triple in taxonomy:
                out.append(CategorisedRow(str(txn.get("transactionId", "")).strip(), owner_slug(txn.get("owner", "")),
                                          *triple, Method.MODEL, round(confidence, 2), "", False, now))
        return out


def train(examples: list[tuple[dict, Triple, float]], taxonomy: Taxonomy, trained_at: str) -> Model:
    """`examples` are (transaction-shaped dict, label, weight). Labels not in the taxonomy,
    and labels with fewer than `MIN_LABEL_EXAMPLES` examples, are left out."""
    import sklearn

    counts = Counter(label for _t, label, _w in examples if label in taxonomy)
    kept = [(t, label, w) for t, label, w in examples if counts[label] >= MIN_LABEL_EXAMPLES]
    if len({label for _t, label, _w in kept}) < 2:
        raise ValueError("Not enough labelled examples to train a model (need two labels with "
                         f"{MIN_LABEL_EXAMPLES}+ examples each).")
    pipeline = _pipeline()
    pipeline.fit(_frame([t for t, _l, _w in kept]), [_SEP.join(label) for _t, label, _w in kept],
                 classifier__sample_weight=[w for _t, _l, w in kept])
    meta = {"format": FORMAT, "trainedAt": trained_at, "sklearn": sklearn.__version__,
            "examples": len(kept), "labels": len({label for _t, label, _w in kept}),
            "labelCounts": {" / ".join(label): n for label, n in counts.most_common() if n >= MIN_LABEL_EXAMPLES}}
    return Model(pipeline, meta)


def save(model: Model, folder: Path) -> Path:
    import joblib

    folder.mkdir(parents=True, exist_ok=True)
    path = folder / MODEL_FILE
    joblib.dump(model.pipeline, path)
    (folder / META_FILE).write_text(json.dumps(model.meta, indent=2) + "\n")
    return path


def load(folder: Path) -> tuple[Model | None, str]:
    """(model, "loaded"), or (None, why): "missing", "unreadable" or "incompatible"
    (other format or scikit-learn version than the one running; retrain)."""
    import sklearn

    path, meta_path = folder / MODEL_FILE, folder / META_FILE
    if not path.is_file() or not meta_path.is_file():
        return None, "missing"
    try:
        meta = json.loads(meta_path.read_text())
    except (OSError, ValueError):
        return None, "unreadable"
    if not isinstance(meta, dict) or meta.get("format") != FORMAT or meta.get("sklearn") != sklearn.__version__:
        return None, "incompatible"
    try:
        import joblib
        pipeline = joblib.load(path)  # our own file in the runtime folder, written by `train`
    except Exception:  # noqa: BLE001 - a broken model file must never stop the daily run
        return None, "unreadable"
    if not hasattr(pipeline, "predict_proba"):
        return None, "unreadable"
    return Model(pipeline, meta), "loaded"


def review_check(rows: list[CategorisedRow], reviewed: dict[str, Triple]) -> tuple[int, int]:
    """(checked, wrong): model predictions for rows you reviewed, and how many disagree."""
    checked = [r for r in rows if r.transactionId in reviewed]
    wrong = sum((r.type, r.category, r.subCategory) != reviewed[r.transactionId] for r in checked)
    return len(checked), wrong
