"""Rules from the `categoriser_rules` sheet: parsing, validation and matching.

A rule matches a transaction when every match column that is filled in matches
(blank means "any"). Text matching is case-insensitive, ignores punctuation and
works on whole words (see `normalise`), so `corner cafe` matches
`Shop 12 - Corner_Cafe Ltd` but `corner caf` does not. `textRegex` is a Python regex searched,
case-insensitively, in the raw "counterparty | description | reference" text.

Rules are tried by `priority` (lowest first), then by how many match columns
are set (most first), then by `ruleId`. The first match wins.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field

from .issues import Issue
from .merge import is_reviewed
from .taxonomy import Taxonomy

RULE_COLUMNS: tuple[str, ...] = (
    "ruleId", "active", "priority",
    # match columns (blank = any)
    "account", "direction", "counterpartyContains", "descriptionContains", "textRegex",
    "providerCategory", "amountMin", "amountMax",
    # outputs (owner blank = the account's owner)
    "owner", "type", "category", "subCategory",
    "notes", "source",
    # written by the job
    "hits", "lastMatched",
)
RCOL = {name: i for i, name in enumerate(RULE_COLUMNS)}
MATCH_COLUMNS = ("account", "direction", "counterpartyContains", "descriptionContains", "textRegex",
                 "providerCategory", "amountMin", "amountMax")
DIRECTIONS = ("in", "out", "any")
SOURCES = ("manual", "mined")
DEFAULT_PRIORITY = 100.0

INVALID_RULE = "invalid_rule"

_NON_WORD = re.compile(r"[\W_]+", re.UNICODE)


def normalise(text: object) -> str:
    """Lower-case words separated by single spaces; punctuation and `_` removed."""
    return _NON_WORD.sub(" ", str(text or "").casefold()).strip()


def contains_words(text: object, needle: object) -> bool:
    """True if the words of `needle` appear together in `text` (whole words, normalised)."""
    words = normalise(needle)
    return bool(words) and f" {words} " in f" {normalise(text)} "


def account_label(txn: dict) -> str:
    return f"{str(txn.get('bank', '')).strip()} / {str(txn.get('accountName', '')).strip()}"


def amount_eur(txn: dict) -> float:
    """Signed EUR amount; falls back to `amount` when `amountEur` is blank."""
    for key in ("amountEur", "amount"):
        value = txn.get(key, "")
        if value not in ("", None):
            try:
                return float(value)
            except (TypeError, ValueError):
                continue
    return 0.0


@dataclass(frozen=True)
class Rule:
    ruleId: str
    active: bool
    priority: float
    account: str = ""
    direction: str = ""
    counterpartyContains: str = ""
    descriptionContains: str = ""
    textRegex: str = ""
    providerCategory: str = ""
    amountMin: float | None = None
    amountMax: float | None = None
    owner: str = ""
    type: str = ""
    category: str = ""
    subCategory: str = ""
    source: str = ""
    _regex: re.Pattern | None = field(default=None, compare=False, repr=False)

    @property
    def specificity(self) -> int:
        return sum(1 for c in MATCH_COLUMNS if getattr(self, c) not in ("", None, "any"))

    def sort_key(self) -> tuple:
        return (self.priority, -self.specificity, self.ruleId)

    def matches(self, txn: dict) -> bool:
        if self.account and normalise(account_label(txn)) != normalise(self.account):
            return False
        amount = amount_eur(txn)
        if self.direction == "in" and amount <= 0:
            return False
        if self.direction == "out" and amount >= 0:
            return False
        if self.counterpartyContains and not contains_words(txn.get("counterparty"), self.counterpartyContains):
            return False
        if self.descriptionContains and not contains_words(txn.get("description"), self.descriptionContains):
            return False
        if self.providerCategory and normalise(self.providerCategory) != normalise(txn.get("category")):
            return False
        if self.amountMin is not None and abs(amount) < self.amountMin:
            return False
        if self.amountMax is not None and abs(amount) > self.amountMax:
            return False
        if self._regex is not None:
            text = " | ".join(str(txn.get(k, "") or "") for k in ("counterparty", "description", "reference"))
            if not self._regex.search(text):
                return False
        return True


@dataclass
class ParsedRules:
    rules: list[Rule]                        # valid rules, active or not, in matching order
    row_of: dict[str, int]                   # ruleId -> 1-based sheet row (valid rules only)
    issues: list[Issue]

    @property
    def active(self) -> list[Rule]:
        return [r for r in self.rules if r.active]


def _number(value: object) -> float | None:
    if value in ("", None):
        return None
    if isinstance(value, bool):
        raise ValueError
    return float(str(value).replace(",", "."))


def _rule_issue(ref: str, detail: str, sheet: str) -> Issue:
    return Issue(type=INVALID_RULE, transactionId="", detail=detail,
                 source=f"categoriser.run: '{sheet}' sheet check", ref=ref)


def parse_rules(values: list[list], taxonomy: Taxonomy, owners: set[str], sheet: str) -> ParsedRules:
    """`values` is the whole rules sheet incl. header, unformatted.

    Invalid rules are left out and reported (active ones only, so a half-written
    inactive rule does not raise an issue). Never raises for bad rule content.
    """
    if values and [str(h).strip() for h in values[0][: len(RULE_COLUMNS)]] != list(RULE_COLUMNS):
        raise ValueError("Header row of the rules sheet does not match the expected columns.")

    def cell(row: list, name: str) -> object:
        i = RCOL[name]
        return row[i] if len(row) > i else ""

    def text(row: list, name: str) -> str:
        return str(cell(row, name) or "").strip()

    rows = [(n, r) for n, r in enumerate(values[1:], start=2) if any(str(c).strip() for c in r)]
    ids = [text(r, "ruleId") for _, r in rows]
    duplicate_ids = {i for i in ids if i and ids.count(i) > 1}

    rules: list[Rule] = []
    row_of: dict[str, int] = {}
    issues: list[Issue] = []
    for row_number, row in rows:
        rule_id = text(row, "ruleId")
        active = is_reviewed(cell(row, "active"))
        problems: list[str] = []
        if not rule_id:
            problems.append("ruleId is empty")
        elif rule_id in duplicate_ids:
            problems.append(f"ruleId '{rule_id}' is used on more than one row")

        try:
            priority = _number(cell(row, "priority"))
            priority = DEFAULT_PRIORITY if priority is None else priority
        except ValueError:
            problems.append("priority is not a number")
            priority = DEFAULT_PRIORITY

        amounts: dict[str, float | None] = {}
        for name in ("amountMin", "amountMax"):
            try:
                amounts[name] = _number(cell(row, name))
            except ValueError:
                problems.append(f"{name} is not a number")
                amounts[name] = None
        if None not in amounts.values() and amounts["amountMin"] > amounts["amountMax"]:
            problems.append("amountMin is larger than amountMax")

        direction = text(row, "direction").lower()
        if direction and direction not in DIRECTIONS:
            problems.append(f"direction must be one of {', '.join(DIRECTIONS)}")

        regex = None
        if text(row, "textRegex"):
            try:
                regex = re.compile(text(row, "textRegex"), re.IGNORECASE)
            except re.error as err:
                problems.append(f"textRegex does not compile ({err})")

        rule = Rule(
            ruleId=rule_id, active=active, priority=priority,
            account=text(row, "account"), direction="" if direction == "any" else direction,
            counterpartyContains=text(row, "counterpartyContains"),
            descriptionContains=text(row, "descriptionContains"),
            textRegex=text(row, "textRegex"), providerCategory=text(row, "providerCategory"),
            amountMin=amounts["amountMin"], amountMax=amounts["amountMax"],
            owner=text(row, "owner"), type=text(row, "type"), category=text(row, "category"),
            subCategory=text(row, "subCategory"), source=text(row, "source").lower(), _regex=regex,
        )
        if rule.specificity == 0 and not problems:
            problems.append("no match column is filled in (it would match every transaction)")
        if (rule.type, rule.category, rule.subCategory) not in taxonomy:
            problems.append(f"type/category/subCategory '{rule.type} / {rule.category} / {rule.subCategory}' "
                            f"is not in the taxonomy sheet")
        if rule.owner and owner_slug(rule.owner) not in owners:
            problems.append(f"owner '{rule.owner}' is not an owner on the transactions sheet")

        if problems:
            if active:
                ref = rule_id or f"row {row_number}"
                issues.append(_rule_issue(ref, f"Rule '{ref}' (row {row_number}) skipped: {'; '.join(problems)}.",
                                          sheet))
            continue
        rules.append(rule)
        row_of[rule_id] = row_number

    rules.sort(key=Rule.sort_key)
    return ParsedRules(rules=rules, row_of=row_of, issues=issues)


def owner_slug(owner: str) -> str:
    """Same as bank-sync's `Connections.slug`, which fills the `owner` column (`Åsa Lind` -> `asa-lind`)."""
    text = unicodedata.normalize("NFD", str(owner))
    text = "".join(c for c in text if not unicodedata.combining(c)).lower()
    return re.sub(r"[^a-z0-9]+", "-", text).strip("-")


def first_match(rules: list[Rule], txn: dict) -> Rule | None:
    for rule in rules:
        if rule.matches(txn):
            return rule
    return None
