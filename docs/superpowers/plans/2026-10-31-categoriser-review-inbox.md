# Categoriser 3a: Review Inbox Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A `categoriser_review` tab that the daily run fills with the rows that most need the owner, and
that turns ticked answers into reviewed `manual` rows (and, with "all like this", active rules) on the next run.

**Architecture:** Pure planning in a new `review.py` (answers → manual rows, rules, issues, history; inbox
rows from transactions and the post-run `categorised` sheet), I/O added to `sheets.py`, wiring in `run.py`.
Answers are applied before categorising so a new rule works in the same run; the inbox is rebuilt after
the `categorised` write.

**Tech Stack:** Python 3.12, gspread, pytest (no network), the in-memory fakes in `tests/fakes.py`.

**Spec:** `docs/superpowers/specs/2026-10-31-categoriser-phase-3-design.md` (section "3a").

All paths below are relative to `bank-sync/categoriser/` unless they start with `docs/` or are the repo's `CLAUDE.md`.
Run tests from `bank-sync/categoriser/` with the venv: `source .venv/bin/activate && pytest -q`.

## Global Constraints

- Never write to `transactions`, the legacy sheet or a row with `reviewed` = TRUE; every write goes through `_assert_writable`.
- Problems are `Issue`s in `categoriser_issues`; nothing in 3a may raise out of `run.main` because of a bad answer.
- Logs carry counts only, never transaction text, amounts or ids.
- Tests use synthetic data only (`tests/builders.py`: alice, bob, Acme Market, Corner Cafe, `XX00TEST…`).
- Settings: `REVIEW_SHEET` = `categoriser_review`, `REVIEW_SIZE` = 25, `REVIEW_SPOT_CHECKS` = 3.
- Answer separator ` › ` (U+203A with a space on each side); review rules: `priority` 50, `source` = `review`, `ruleId` = `r-` + 8 hex chars.
- Issue types: `invalid_review_answer`, `review_conflicts_with_reviewed`.
- Review history: one JSON object per line in `<CATEGORISER_HOME>/data/review-history.jsonl`, never in the repo.
- After merge, deploying = re-running `bank-sync/categoriser/scripts/install-schedule.sh` from `main`.

## Review Focus

1. **The owner sorts or filters the review tab on the phone.** Rows are found by `transactionId`, never by
   position; a sorted tab must apply the same answers. → test in Task 2 (`test_row_order_does_not_matter`).
2. **The same transaction appears twice in the inbox** (owner copied a row). The first ticked copy wins, the
   second is dropped without a second manual row or a crash in `plan_upsert`. → Task 2 (`test_duplicate_inbox_rows_apply_once`).
3. **An answer is applied but the inbox rewrite is skipped (busy).** Next run the same ticked row comes back;
   it must not write a second history line or a second rule. → Task 5 (`test_answer_seen_twice_is_applied_once`).
4. **Checkbox cells read back as the string `"TRUE"`/`"FALSE"` or as blank** (old rows, typed values). Use
   `merge.is_reviewed` for both checkboxes. → Task 2 (`test_ticks_as_text_count`).
5. **`matchText` the owner shortened to a word with punctuation or capitals** (`"Corner-Cafe"`). It is
   normalised before it becomes the rule and the id. → Task 2 (`test_match_text_is_normalised`).

---

### Task 1: Settings, review layout and helpers

**Files:**
- Modify: `src/categoriser/config.py` (Settings fields, `writable_sheets`, `load_settings`, new `data_dir()`)
- Modify: `src/categoriser/rules.py:38` (`SOURCES`)
- Create: `src/categoriser/review.py` (layout and helpers only in this task)
- Modify: `tests/builders.py` (add `review_row`, `review_sheet`)
- Test: `tests/test_review_layout.py`, `tests/test_sheets_guard.py`

**Interfaces:**
- Produces:
  - `Settings.review_sheet: str = "categoriser_review"`, `Settings.review_size: int = 25`,
    `Settings.review_spot_checks: int = 3`; review sheet in `writable_sheets` and in the must-differ check;
    `__post_init__` raises `ValueError` unless `review_size >= 1` and `0 <= review_spot_checks <= review_size`.
    `load_settings` reads `REVIEW_SHEET`, `REVIEW_SIZE`, `REVIEW_SPOT_CHECKS` (an `_int` helper like `_float`).
  - `config.data_dir() -> Path` = `categoriser_home() / "data"`.
  - `rules.SOURCES = ("manual", "mined", "review")`.
  - In `review.py`: `REVIEW_COLUMNS = ("transactionId", "bookingDate", "account", "amountEur", "counterparty",
    "description", "why", "similar", "confidence", "answer", "owner", "accept", "allLikeThis", "matchText",
    "suggested", "kind")`, `VCOL` (name → index), `ANSWER_SEP = " › "`, `KIND_MERCHANT = "merchant"`,
    `KIND_SPOT_CHECK = "spot check"`, `REVIEW_RULE_PRIORITY = 50`, `INVALID_REVIEW = "invalid_review_answer"`,
    `REVIEW_CONFLICT = "review_conflicts_with_reviewed"`.
  - `answer_text(triple: Triple) -> str`, `parse_answer(text: object) -> Triple | None`.
  - `merchant_key(txn: dict) -> tuple[str, str, str]` = `(field, words, direction)`.
  - `review_rule_id(field: str, words: str, direction: str) -> str`.
  - `ticked(review: list[list]) -> set[str]`: transactionIds of rows with `accept` TRUE (header skipped).
  - `builders.review_row(tid, answer="", owner="", accept=False, all_like_this=False, match_text="",
    suggested="", kind="merchant", confidence="") -> list` and `builders.review_sheet(*rows) -> list[list]` (header first).

- [ ] **Step 1: Write the failing tests** in `tests/test_review_layout.py`:

```python
def test_answer_round_trips_and_keeps_slashes():
    t = ("Wants/Pleasure", "Eating out", "Cafe")
    assert answer_text(t) == "Wants/Pleasure › Eating out › Cafe"
    assert parse_answer("  Wants/Pleasure ›Eating out ›  Cafe ") == t
    assert parse_answer("") is None and parse_answer(None) is None
    assert parse_answer("Essentials › Groceries") is None

def test_merchant_key_is_the_longest_run_of_words_without_numbers():
    assert merchant_key(txn(counterparty="Corner_Cafe 1234 Ltd", amount=-3)) == ("counterpartyContains", "corner cafe", "out")
    assert merchant_key(txn(counterparty="PAYPAL *ACME 4029357733")) == ("counterpartyContains", "paypal acme", "out")

def test_merchant_key_falls_back_to_the_description():
    assert merchant_key(txn(counterparty="12345", description="SEPA Rent 2026", amount=900)) == ("descriptionContains", "sepa rent", "in")
    assert merchant_key(txn(counterparty="", description="")) == ("descriptionContains", "", "out")

def test_merchant_key_words_always_match_their_row():
    t = txn(counterparty="Shop 12 - Corner_Cafe Ltd")
    assert contains_words(t["counterparty"], merchant_key(t)[1])

def test_review_rule_id_is_stable_and_normalised():
    a = review_rule_id("counterpartyContains", "Corner-Cafe", "out")
    assert a == review_rule_id("counterpartyContains", "corner cafe", "out") and a.startswith("r-") and len(a) == 10
    assert a != review_rule_id("counterpartyContains", "corner cafe", "in")

def test_ticked_reads_checkboxes_and_text():
    sheet = review_sheet(review_row("a", accept=True), review_row("b", accept="TRUE"), review_row("c"))
    assert ticked(sheet) == {"a", "b"} and ticked([]) == set()
```

In `tests/test_sheets_guard.py`, add `"categoriser_review"` to the allowed-titles parametrisation, add
`("review_sheet", "transactions")` to the collision cases, and add
`test_review_size_and_spot_checks_must_fit` (parametrised: `review_size=0`; `review_spot_checks=-1`;
`review_size=3, review_spot_checks=4` → `ValueError`).

- [ ] **Step 2: Run** `pytest -q tests/test_review_layout.py tests/test_sheets_guard.py` → FAIL (`ModuleNotFoundError: categoriser.review`).
- [ ] **Step 3: Implement** the Interfaces above. `merchant_key`: `normalise` the counterparty, split into words,
  take the longest run of consecutive words containing no digit (first on a tie); if empty, the same on the
  description with field `descriptionContains`. Direction is `"in"` when `rules.amount_eur(txn) > 0`, else `"out"`.
  `review_rule_id` hashes `f"{field}|{normalise(words)}|{direction}"` with sha1, like `MinedRule.rule_id`.
  `parse_answer` splits on `›`, strips each part, needs exactly three non-empty parts.
- [ ] **Step 4: Run** the same command → PASS; then `pytest -q` → all pass.
- [ ] **Step 5: Commit** `feat(categoriser): review tab layout and settings`.

---

### Task 2: Applying answers

**Files:**
- Modify: `src/categoriser/review.py`
- Test: `tests/test_review_answers.py`

**Interfaces:**
- Consumes: Task 1 helpers; `schema.CategorisedRow`, `Method.MANUAL`; `rules.RULE_COLUMNS`, `RCOL`,
  `owner_slug`, `normalise`, `amount_eur`; `merge.is_reviewed`; `issues.Issue`; `taxonomy.Taxonomy`.
- Produces:

```python
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

def apply_answers(review: list[list], transactions: dict[str, dict], taxonomy: Taxonomy, owners: set[str],
                  categorised: list[list], rule_ids: set[str], today: str, now: str, sheet: str) -> AnswerPlan
```

`review` and `categorised` are whole sheets incl. header; `transactions` is by transactionId; `rule_ids` are
the ruleIds already on the rules sheet; `today` is `YYYY-MM-DD`, `now` ISO UTC.

Per ticked row, in sheet order (unticked rows are ignored; a transactionId already handled this run is skipped):

| condition | result |
|---|---|
| id not in `transactions` | keep + `invalid_review_answer` ("not on transactions") |
| `parse_answer` is None or triple not in taxonomy | keep + `invalid_review_answer` |
| owner (`owner_slug` of the cell, or of the account's owner when blank) not in `owners` | keep + `invalid_review_answer` |
| id on more than one row of `categorised` | keep + `invalid_review_answer` ("on several rows of categorised") |
| reviewed row on `categorised`, same owner and triple | processed only |
| reviewed row on `categorised`, different | keep + `review_conflicts_with_reviewed` |
| otherwise | manual row + history + processed (+ rule, below) |

Issues: `Issue(type, transactionId=tid, detail=..., source=f"categoriser.run: '{sheet}' answers")`.
The detail names the problem, never amounts or text from the transaction.

"All like this" (ticked and `normalise(matchText)` not blank): `field` = `merchant_key(txn)[0]`, direction from
the amount, `rid = review_rule_id(field, matchText, direction)`. The manual row's `ruleId` = `rid`. A rule row is
added unless `rid` is in `rule_ids` or already added this run: `ruleId`=rid, `active`=True, `priority`=50,
`direction`, `<field>` = `normalise(matchText)`, `owner` = the answer's owner or `""` when it equals the
account's owner slug, `type`/`category`/`subCategory`, `notes` = `f"from review {today}"`, `source` = `"review"`,
everything else `""`.

History entry: `{"date": today, "transactionId": tid, "suggested": <suggested cell>, "confidence": <float or None>,
"answer": answer_text(triple), "spotCheck": kind == "spot check", "card": classification == "card"}`.

- [ ] **Step 1: Write the failing tests** in `tests/test_review_answers.py` (use `builders.txn`, `TAXONOMY`,
  `OWNERS`, `review_row`, `review_sheet`; `HEADER = list(CATEGORISED_COLUMNS)`; a helper
  `plan(review, txns, categorised=[HEADER], rule_ids=set())` calling `apply_answers` with `today="2026-11-01"`,
  `now="2026-11-01T06:15:00Z"`, `sheet="categoriser_review"`):

```python
GROCERY = "Essentials › Groceries › Supermarket"

def test_accepted_answer_becomes_a_reviewed_manual_row():
    p = plan(review_sheet(review_row("a", answer=GROCERY, accept=True)), [txn("bob", tid="a")])
    [row] = p.manual_rows
    assert (row.transactionId, row.owner, row.type, row.method, row.reviewed, row.confidence, row.ruleId) == \
        ("a", "bob", "Essentials", Method.MANUAL, True, None, "")
    assert p.processed == {"a"} and p.keep == [] and p.issues == []

def test_unticked_rows_are_ignored():
    p = plan(review_sheet(review_row("a", answer=GROCERY)), [txn(tid="a")])
    assert p.manual_rows == [] and p.processed == set()

@pytest.mark.parametrize("answer,owner,tid", [
    ("Essentials › Groceries › Bakery", "", "a"),   # not in the taxonomy
    ("", "", "a"),                                   # no answer
    (GROCERY, "carol", "a"),                         # unknown owner
    (GROCERY, "", "gone"),                           # not on transactions
])
def test_invalid_answers_stay_ticked_with_an_issue(answer, owner, tid):
    row = review_row(tid, answer=answer, owner=owner, accept=True)
    p = plan(review_sheet(row), [txn(tid="a")])
    assert p.manual_rows == [] and p.keep == [row] and p.processed == set()
    assert [i.type for i in p.issues] == ["invalid_review_answer"]

def test_id_on_several_categorised_rows_is_not_applied():
    dup = ["a", "", "", "", "", "", "", "", False, ""]
    p = plan(review_sheet(review_row("a", answer=GROCERY, accept=True)), [txn(tid="a")], [HEADER, dup, dup])
    assert p.manual_rows == [] and [i.type for i in p.issues] == ["invalid_review_answer"]

def test_answer_matching_a_reviewed_row_is_done_without_writing():
    reviewed = ["a", "alice", "Essentials", "Groceries", "Supermarket", "manual", "", "", True, ""]
    p = plan(review_sheet(review_row("a", answer=GROCERY, accept=True)), [txn("alice", tid="a")], [HEADER, reviewed])
    assert p.manual_rows == [] and p.history == [] and p.processed == {"a"}

def test_answer_conflicting_with_a_reviewed_row_stays_with_an_issue():
    reviewed = ["a", "alice", "Wants/Pleasure", "Eating out", "Cafe", "manual", "", "", True, ""]
    row = review_row("a", answer=GROCERY, accept=True)
    p = plan(review_sheet(row), [txn("alice", tid="a")], [HEADER, reviewed])
    assert p.keep == [row] and [i.type for i in p.issues] == ["review_conflicts_with_reviewed"]

def test_all_like_this_adds_one_valid_active_rule():
    t1, t2 = txn("bob", counterparty="Acme Market 77", tid="a"), txn("bob", counterparty="Acme Market", tid="b")
    p = plan(review_sheet(review_row("a", answer=GROCERY, accept=True, all_like_this=True, match_text="acme market"),
                          review_row("b", answer=GROCERY, accept=True, all_like_this=True, match_text="acme market")),
             [t1, t2])
    [rule_values] = p.new_rules
    parsed = parse_rules(rules_sheet(rule_values), TAXONOMY, OWNERS, "categoriser_rules")
    [rule] = parsed.active
    assert parsed.issues == [] and rule.matches(t1) and rule.matches(t2)
    assert (rule.priority, rule.source, rule.direction, rule.owner) == (50, "review", "out", "")
    assert rule_values[RCOL["notes"]] == "from review 2026-11-01"
    assert {r.ruleId for r in p.manual_rows} == {rule.ruleId}

def test_existing_rule_id_is_not_added_again():
    rid = review_rule_id("counterpartyContains", "acme market", "out")
    p = plan(review_sheet(review_row("a", answer=GROCERY, accept=True, all_like_this=True, match_text="acme market")),
             [txn(counterparty="Acme Market", tid="a")], rule_ids={rid})
    assert p.new_rules == [] and p.manual_rows[0].ruleId == rid

def test_rule_keeps_an_owner_other_than_the_accounts():
    p = plan(review_sheet(review_row("a", answer=GROCERY, owner="common", accept=True, all_like_this=True,
                                     match_text="acme market")), [txn("bob", counterparty="Acme Market", tid="a")])
    assert p.new_rules[0][RCOL["owner"]] == "common" and p.manual_rows[0].owner == "common"

def test_history_and_counts():
    cafe = "Wants/Pleasure › Eating out › Cafe"
    p = plan(review_sheet(review_row("a", answer=GROCERY, suggested=cafe, accept=True),
                          review_row("b", answer=cafe, suggested=cafe, accept=True, kind="spot check", confidence=0.96),
                          review_row("c", answer=GROCERY, suggested=cafe, accept=True, kind="spot check")),
             [txn(tid="a", classification="card"), txn(tid="b"), txn(tid="c")])
    assert (p.changed, p.spot_right, p.spot_wrong) == (2, 1, 1)
    assert p.history[0] == {"date": "2026-11-01", "transactionId": "a", "suggested": cafe, "confidence": None,
                            "answer": GROCERY, "spotCheck": False, "card": True}
    assert p.history[1]["confidence"] == 0.96
```

Plus the Review Focus tests:

```python
def test_row_order_does_not_matter():          # same two answers, rows reversed → same manual rows (by id)
def test_duplicate_inbox_rows_apply_once():     # two ticked rows for "a" → one manual row, one history entry, no issue
def test_ticks_as_text_count():                 # accept="TRUE" applies; accept="FALSE" and "" do not
def test_match_text_is_normalised():            # match_text="Corner-Cafe" → rule counterpartyContains == "corner cafe",
                                                # ruleId == review_rule_id("counterpartyContains", "corner cafe", "out")
```

- [ ] **Step 2: Run** `pytest -q tests/test_review_answers.py` → FAIL (`ImportError: apply_answers`).
- [ ] **Step 3: Implement** `AnswerPlan` and `apply_answers` per the table above. Read cells by `VCOL`, padding
  short rows; read `categorised` by `CATEGORISED_COLUMNS` (reviewed rows and row counts per id).
- [ ] **Step 4: Run** `pytest -q tests/test_review_answers.py` → PASS; `pytest -q` → all pass.
- [ ] **Step 5: Commit** `feat(categoriser): apply review answers`.

---

### Task 3: Filling the inbox

**Files:**
- Modify: `src/categoriser/review.py`
- Test: `tests/test_review_inbox.py`

**Interfaces:**
- Consumes: Task 1 helpers, `AnswerPlan` (Task 2).
- Produces:

```python
Guess = Callable[[list[dict]], list[tuple[Triple, float]]]   # Model.predict has this shape

def plan_inbox(previous: list[list], answers: AnswerPlan, transactions: list[dict], labelled: set[str],
               model_ids: set[str], taxonomy: Taxonomy, guess: Guess | None, size: int, spot_checks: int,
               rng: random.Random) -> list[list]
```

`previous` = the review tab as read at the start (incl. header). `labelled` = transactionIds with a full triple
on `categorised` after this run's write. `model_ids` = ids the model labelled this run. Returns the new inbox
rows **without** header, each `len(REVIEW_COLUMNS)` long.

Algorithm:
1. **Kept** = `answers.keep`, then each unticked previous row (first copy per id, not in `answers.processed`,
   id still on `transactions`) when `kind` is `merchant` and the id is not in `labelled`, or `kind` is
   `spot check` and the id is in `model_ids`. Kept rows are copied unchanged. Kept rows are never dropped,
   even beyond `size`.
2. **Candidates** = transactions (first copy per non-blank id) not in `labelled`, not kept, grouped by
   `merchant_key`; groups whose key a kept row's transaction has are skipped.
3. `spot_new = max(0, spot_checks - kept spot checks)`; `merchant_new = max(0, size - len(kept) - spot_new)`.
4. Take `merchant_new` groups by size (largest first, then key), one row each: the latest `bookingDate`, then
   the largest id. `why` = `"{n} rows like this have no category"` (`"1 row like this has no category"` for 1),
   `similar` = n, `matchText` = the key's words, `kind` = `merchant`.
5. Spot checks: `rng.sample` of `spot_new` ids from sorted `model_ids` minus kept ids and minus ids whose
   `merchant_key` is in a chosen or kept group (fewer if there aren't enough). `why` =
   `"spot check: model {pct}% sure"`, `similar` = `""`, `matchText` = key words, `kind` = `spot check`.
6. Pre-fill new rows: one `guess` call with all new rows' transactions (card rows included). A guessed triple
   in the taxonomy fills `answer` and `suggested` with `answer_text(triple)` and `confidence` with
   `round(p, 2)`; otherwise these three stay `""`. `owner` = `owner_slug(txn["owner"])`, `account` =
   `rules.account_label`, `amountEur` = `rules.amount_eur`, `bookingDate` = first 10 chars, `accept` and
   `allLikeThis` = False.
7. Order: kept and new merchant rows by `similar` (largest first; blank counts as 0), then all spot checks.

- [ ] **Step 1: Write the failing tests** in `tests/test_review_inbox.py` (helper
  `inbox(txns, previous=None, answers=None, labelled=(), model_ids=(), guess=None, size=25, spot=0)`
  passing `random.Random(1)` and `TAXONOMY`):

```python
def test_biggest_merchant_groups_first_one_latest_row_each():
    txns = [txn(counterparty="Acme Market", day=f"2026-05-0{d}", tid=f"a{d}") for d in (1, 2, 3)]
    txns += [txn(counterparty="Corner Cafe", day=f"2026-05-0{d}", tid=f"c{d}") for d in (1, 2)]
    txns += [txn(counterparty="Lone Shop", tid="l")]
    rows = inbox(txns, size=2)
    assert [r[VCOL["transactionId"]] for r in rows] == ["a3", "c2"]
    assert rows[0][VCOL["why"]] == "3 rows like this have no category" and rows[0][VCOL["similar"]] == 3
    assert rows[0][VCOL["matchText"]] == "acme market" and rows[0][VCOL["kind"]] == "merchant"

def test_labelled_rows_are_not_candidates():
    assert inbox([txn(counterparty="Acme Market", tid="a")], labelled={"a"}) == []

def test_unanswered_row_survives_with_its_typed_answer_and_its_group_is_not_added_again():
    old = review_row("a1", answer="Essentials › Groceries › Supermarket")
    txns = [txn(counterparty="Acme Market", tid="a1"), txn(counterparty="Acme Market", tid="a2")]
    rows = inbox(txns, previous=review_sheet(old))
    assert rows == [old + [""] * (len(REVIEW_COLUMNS) - len(old))]

def test_rows_that_got_a_label_leave_the_inbox():
    rows = inbox([txn(counterparty="Acme Market", tid="a1")], previous=review_sheet(review_row("a1")), labelled={"a1"})
    assert rows == []

def test_ticked_rows_that_were_not_applied_stay():
    bad = review_row("a1", answer="nope", accept=True)
    rows = inbox([txn(tid="a1")], previous=review_sheet(bad), answers=AnswerPlan(keep=[bad]), labelled={"a1"})
    assert [r[VCOL["transactionId"]] for r in rows] == ["a1"]

def test_spot_checks_take_the_last_slots_from_model_rows():
    txns = [txn(counterparty="Acme Market", tid="a"), txn(counterparty="Corner Cafe", tid="m1"),
            txn(counterparty="Lone Shop", tid="m2")]
    rows = inbox(txns, labelled={"m1", "m2"}, model_ids={"m1", "m2"}, size=2, spot=1,
                 guess=lambda ts: [(("Essentials", "Groceries", "Supermarket"), 0.96)] * len(ts))
    assert [r[VCOL["kind"]] for r in rows] == ["merchant", "spot check"]
    assert rows[1][VCOL["why"]] == "spot check: model 96% sure" and rows[1][VCOL["transactionId"]] in {"m1", "m2"}

def test_prefill_uses_the_model_guess_only_inside_the_taxonomy():
    txns = [txn(counterparty="Acme Market", tid="a", classification="card"), txn(counterparty="Odd", tid="o")]
    guesses = {"a": (("Essentials", "Groceries", "Supermarket"), 0.734), "o": (("No", "Such", "Label"), 0.99)}
    rows = {r[0]: r for r in inbox(txns, guess=lambda ts: [guesses[t["transactionId"]] for t in ts])}
    assert rows["a"][VCOL["answer"]] == rows["a"][VCOL["suggested"]] == "Essentials › Groceries › Supermarket"
    assert rows["a"][VCOL["confidence"]] == 0.73 and rows["o"][VCOL["answer"]] == rows["o"][VCOL["confidence"]] == ""

def test_without_a_model_the_answer_is_blank_and_owner_is_prefilled():
    [row] = inbox([txn("bob", counterparty="Acme Market", tid="a")])
    assert row[VCOL["answer"]] == "" and row[VCOL["owner"]] == "bob" and row[VCOL["accept"]] is False

def test_kept_rows_are_never_dropped_and_no_new_rows_beyond_size():
    olds = [review_row(f"k{i}") for i in range(3)]
    txns = [txn(counterparty=f"Shop {chr(97 + i)}", tid=f"k{i}") for i in range(3)] + [txn(counterparty="New", tid="n")]
    assert len(inbox(txns, previous=review_sheet(*olds), size=2)) == 3
```

- [ ] **Step 2: Run** `pytest -q tests/test_review_inbox.py` → FAIL (`ImportError: plan_inbox`).
- [ ] **Step 3: Implement** `plan_inbox` per the algorithm.
- [ ] **Step 4: Run** `pytest -q tests/test_review_inbox.py` → PASS; `pytest -q` → all pass.
- [ ] **Step 5: Commit** `feat(categoriser): fill the review inbox`.

---

### Task 4: Review tab I/O

**Files:**
- Modify: `src/categoriser/sheets.py`
- Test: `tests/test_sheets_io.py`

**Interfaces:**
- Consumes: `review.REVIEW_COLUMNS`, `VCOL`, `ticked` (Task 1).
- Produces:
  - `ensure_review_sheet(ss, settings) -> gspread.Worksheet` (via `_ensure_sheet`).
  - `_validations` for the review sheet: `accept` and `allLikeThis` → `CHECKBOX`; `owner` →
    `_one_of_range(settings.owners_sheet, "A")`.
  - `write_review(ws, settings, rows: list[list], ticked_at_read: set[str], answers: list[str]) -> bool`:
    `_assert_writable`; `batch_get` the `transactionId` and `accept` columns (as `upsert_categorised` does);
    if `ticked(...)` of that fresh read differs from `ticked_at_read`, write nothing and return `False`.
    Otherwise write header + rows, then fully blank rows (`""` in every cell) down to the old last row, in one
    `ws.update(..., "A1", value_input_option=RAW)` (add rows first if the grid is too small), re-apply
    `_apply_validations`, and set the `answer` column (row 2 down) to `_one_of_list(tuple(answers), strict=False)`
    when `answers` is non-empty. Return `True`.
  - `_one_of_list(values, strict: bool = True)`.
  - `append_rows` also calls `_apply_validations(ws.spreadsheet, ws, settings)` after writing, so an existing
    rules sheet's `source` dropdown gains `review`.

- [ ] **Step 1: Write the failing tests** in `tests/test_sheets_io.py`:

```python
def test_creates_review_sheet_with_checkboxes_and_owner_dropdown():
    # ensure_review_sheet on an empty FakeSpreadsheet: header == list(REVIEW_COLUMNS), frozen_rows == 1,
    # requests set BOOLEAN on columns VCOL["accept"], VCOL["allLikeThis"] and ONE_OF_RANGE "='account_owners'!A2:A" on VCOL["owner"]

def test_write_review_replaces_rows_and_blanks_leftovers():
    # sheet with header + 3 old rows; write_review(ws, settings, [row_x], set(), ["A › B › C"]) is True;
    # values[1] == row_x, values[2] and values[3] are all "", len(values) == 4

def test_write_review_skips_when_a_row_was_ticked_meanwhile():
    # ws.before_batch_get ticks row 2's accept cell; write_review(..., ticked_at_read=set(), ...) is False
    # and ws.values equals the values right after the tick

def test_answer_dropdown_lists_answers_and_is_not_strict():
    # after write_review with answers ["A › B › C", "D › E › F"], the last setDataValidation request on
    # VCOL["answer"] is ONE_OF_LIST with those two values and "strict": False

def test_write_review_refuses_other_sheets():
    # FakeWorksheet("transactions") → PermissionError

def test_append_rows_reapplies_rule_dropdowns_with_review_source():
    # append_rows on an existing rules sheet → a ONE_OF_LIST request on RCOL["source"] containing "review"
```

Write each with real asserts on the values in the comments.

- [ ] **Step 2: Run** `pytest -q tests/test_sheets_io.py` → the six new tests FAIL.
- [ ] **Step 3: Implement** the Interfaces above.
- [ ] **Step 4: Run** `pytest -q` → all pass (existing `append_rows` tests still pass with the extra request).
- [ ] **Step 5: Commit** `feat(categoriser): read and write the review tab`.

---

### Task 5: Wire into the daily run, history file and docs

**Files:**
- Modify: `src/categoriser/run.py`
- Modify: `tests/test_run.py`
- Modify: `README.md`, `.env.example`, repo `CLAUDE.md` (categoriser section), `docs/BANK_SYNC_ROADMAP.md`
  (remaining-work item 3: "phase 3a review inbox merged"; Last updated line)

**Interfaces:**
- Consumes: everything above; `config.data_dir`, `Model.predict`.
- Produces: `run.main()` order:
  1. `ensure_*` for all sheets incl. `ensure_review_sheet`; read transactions, taxonomy, rules values,
     `categorised` (`before_upsert`, moved up) and the review tab; `owners`.
  2. `answers = apply_answers(review_values, {id: txn}, taxonomy, owners, before_upsert, {rule ids}, today, now,
     settings.review_sheet)`; if `answers.new_rules`: `append_rows(rules_ws, settings, answers.new_rules)` and
     re-read the rules values.
  3. `parse_rules`, `categorise`, model rows as today.
  4. Drop result and model rows whose id is in `{r.transactionId for r in answers.manual_rows}`; upsert
     `result.rows + model_rows + answers.manual_rows`.
  5. `_append_history(data_dir() / "review-history.jsonl", answers.history)` (creates the folder; one
     `json.dumps(entry, ensure_ascii=False)` per line; nothing written for an empty list).
  6. Re-read `categorised`; `labelled` = ids with type, category and subCategory all set. `rows = plan_inbox(
     review_values, answers, transactions, labelled, {r.transactionId for r in model_rows}, taxonomy,
     model.predict if model else None, settings.review_size, settings.review_spot_checks, random.Random(today))`;
     `written = write_review(review_ws, settings, rows, ticked(review_values), [answer_text(t) for t in
     sorted(taxonomy.triples)])`.
  7. Issues: add `answers.issues`; `CHECKED_TYPES` gains `INVALID_REVIEW` and `REVIEW_CONFLICT`.
  8. Log line gains `review_answers=%d review_invalid=%d review_rules_added=%d review_changed=%d
     review_spot_checks_right=%d review_spot_checks_wrong=%d review_size=%d review=%s`
     (`review_answers` = manual rows, `review_invalid` = `len(answers.keep)`, `review` = `written` or `skipped_busy`).

- [ ] **Step 1: Write the failing tests** in `tests/test_run.py`. Add an autouse fixture patching
  `run.data_dir` to `tmp_path / "data"`; extend `spreadsheet(...)` with an optional `review` sheet.

```python
def test_answer_with_all_like_this_labels_the_merchant_in_the_same_run(monkeypatch, tmp_path, caplog):
    # transactions: a1, a2 (bob, Acme Market), n (Lone Shop); review: a1 ticked, answer GROCERY, all_like_this, "acme market"
    # after one run: categorised a1 = manual, reviewed True, ruleId r-…; a2 = rule with the same ruleId;
    # categoriser_rules has one row with source "review"; review tab rows == ["n"];
    # data/review-history.jsonl has 1 line with transactionId "a1"; "review_answers=1" and "review=written" in the log

def test_second_run_changes_nothing(monkeypatch):
    # same setup, run twice: categorised, rules and review values identical after run 1 and run 2

def test_answer_seen_twice_is_applied_once(monkeypatch, tmp_path):
    # run 1 with the review ws's before_batch_get hook ticking another row → "review=skipped_busy";
    # remove the hook, run 2: still one rule row, one history line, a1 reviewed manual

def test_invalid_answer_is_logged_and_stays_ticked(monkeypatch):
    # answer "Essentials › Groceries › Bakery" ticked → issues has "invalid_review_answer:a1" open,
    # review tab still has a1 with accept True, categorised has no a1

def test_inbox_is_created_and_filled_on_the_first_run(monkeypatch):
    # no review sheet: after the run, categoriser_review exists, header == REVIEW_COLUMNS, uncovered rows listed
```

Write each with real asserts on the values in the comments. Existing run tests must still pass unchanged.

- [ ] **Step 2: Run** `pytest -q tests/test_run.py` → the new tests FAIL.
- [ ] **Step 3: Implement** the `run.main()` order above.
- [ ] **Step 4: Run** `pytest -q` → all pass.
- [ ] **Step 5: Docs.** README: a "Review inbox (`categoriser_review`)" section (columns table, what ticking
  does, "all like this", busy skip, history file), the output contract's writable list, the new issue types,
  `source` = `review` in the rules table, `review.py` in the layout. `.env.example`: `REVIEW_SHEET`,
  `REVIEW_SIZE`, `REVIEW_SPOT_CHECKS` with one comment line each. CLAUDE.md: add `categoriser_review` to the
  sheets the job writes and say answers there become reviewed `manual` rows. Roadmap: item 3 notes phase 3a.
- [ ] **Step 6: Local check against the real sheet** (counts only, no data in the PR): `docker compose run
  --rm --build categoriser` twice from `bank-sync/categoriser/`; first run creates the tab with 25 rows,
  second run logs `updated=0 appended=0` and `review=written`.
- [ ] **Step 7: Commit** `feat(categoriser): review inbox in the daily run`, push, open a **draft** PR titled
  `feat(categoriser): review inbox (phase 3a)` with summary and test plan (incl. "After merge: re-run
  `install-schedule.sh`"); mark ready when CI is green.
