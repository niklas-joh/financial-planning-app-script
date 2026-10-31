"""Synthetic test data only: made-up people (alice, bob), merchants and IBANs."""

from __future__ import annotations

from categoriser.rules import RULE_COLUMNS
from categoriser.taxonomy import Taxonomy

TAXONOMY = Taxonomy(frozenset({
    ("Essentials", "Groceries", "Supermarket"),
    ("Wants/Pleasure", "Eating out", "Lunch"),
    ("Wants/Pleasure", "Eating out", "Cafe"),
    ("Income", "Other", "Cashback"),
    ("Savings", "Savings", "Investments"),
    ("Internal Transfer", "Internal Transfer", "Internal Transfer"),
    ("Internal Transfer", "Internal Transfer", "From Alice"),
    ("Internal Transfer", "Internal Transfer", "From Bob"),
}))
OWNERS = {"alice", "bob", "common"}

ACCOUNTS = {
    "alice": dict(bank="Testbank", accountName="Alice Current", accountId="acc-a", ibanMasked="XX00TEST0001", owner="alice"),
    "bob": dict(bank="Testbank", accountName="Bob Current", accountId="acc-b", ibanMasked="XX00TEST0002", owner="bob"),
    "common": dict(bank="Testbank", accountName="Shared Current", accountId="acc-c", ibanMasked="XX00TEST0003", owner="common"),
    "card": dict(bank="Cardco", accountName="Card EUR", accountId="acc-d", ibanMasked="", owner="alice"),
}

_counter = iter(range(1, 100000))


def txn(account="alice", amount=-10.0, counterparty="", description="", day="2026-05-04", tid=None,
        category="", classification="", reference="", counterpartyAccount="", currency="EUR", **extra) -> dict:
    t = {"transactionId": tid or f"t{next(_counter)}", "bookingDate": day, "amount": amount, "amountEur": amount,
         "currency": currency, "counterparty": counterparty, "description": description, "category": category,
         "classification": classification, "reference": reference, "counterpartyAccount": counterpartyAccount}
    t.update(ACCOUNTS[account])
    t.update(extra)
    return t


def rule_row(rule_id="R1", active=True, priority="", output=("Essentials", "Groceries", "Supermarket"),
             source="manual", **match) -> list:
    values = {c: "" for c in RULE_COLUMNS}
    values.update(ruleId=rule_id, active=active, priority=priority, source=source)
    values["type"], values["category"], values["subCategory"] = output
    values.update(match)
    return [values[c] for c in RULE_COLUMNS]


def rules_sheet(*rows) -> list[list]:
    return [list(RULE_COLUMNS), *rows]
