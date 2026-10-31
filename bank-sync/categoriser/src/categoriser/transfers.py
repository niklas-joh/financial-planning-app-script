"""Internal transfers between the household's own accounts.

Every row on `transactions` belongs to one of our accounts, so money leaving one
account and arriving in another is a transfer, not spending or income. Two ways
to spot one:

1. A pair: an outgoing and an incoming row on two different accounts, same
   currency, same amount to the cent, booked at most `MAX_DAYS` apart. Card rows
   (`classification` = card) are never paired. Pairs are formed one to one:
   first those where one leg's `counterpartyAccount` is the other's IBAN, then
   those with the same description (the transfer's message), then by closest
   dates. So two people sending the same amount on the same day stay apart.
2. One leg on its own: its `counterpartyAccount` is the IBAN of another of our
   accounts (the other leg may not be imported).

Pairs are found among all rows, but only legs no rule covered are labelled as
transfers, so a rule can still say that, say, a monthly transfer to a broker is
Savings.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime

from .rules import normalise, owner_slug

MAX_DAYS = 3
PAIR, OWN_ACCOUNT = "transfer:pair", "transfer:own-account"


@dataclass(frozen=True)
class Transfer:
    transactionId: str
    ruleId: str          # PAIR or OWN_ACCOUNT
    owners: frozenset[str]  # owner slugs of both sides (one if both legs are the same person)


def _day(value: object) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    try:
        return date.fromisoformat(str(value).strip()[:10])
    except ValueError:
        return None


def _cents(value: object) -> int | None:
    try:
        return round(float(value) * 100)
    except (TypeError, ValueError):
        return None


def _iban(value: object) -> str:
    return "".join(str(value or "").split()).upper()


def _names_other(a: dict, b: dict) -> bool:
    """Either leg's counterpartyAccount is the other leg's own IBAN."""
    return any(_iban(x.get("counterpartyAccount")) and _iban(x.get("counterpartyAccount")) == _iban(y.get("ibanMasked"))
               for x, y in ((a, b), (b, a)))


def find_transfers(transactions: list[dict]) -> dict[str, Transfer]:
    """transactionId -> Transfer for every row that is an internal transfer."""
    rows = []
    for t in transactions:
        tid = str(t.get("transactionId", "")).strip()
        cents, day = _cents(t.get("amount")), _day(t.get("bookingDate"))
        if tid and cents and day and str(t.get("classification", "")).strip().lower() != "card":
            rows.append((tid, cents, day, t))

    by_key: dict[tuple, list] = {}
    for row in rows:
        if row[1] > 0:
            by_key.setdefault((row[1], str(row[3].get("currency", "")).upper()), []).append(row)

    candidates = []
    for out in rows:
        if out[1] >= 0:
            continue
        for inc in by_key.get((-out[1], str(out[3].get("currency", "")).upper()), []):
            gap = abs((inc[2] - out[2]).days)
            if gap <= MAX_DAYS and str(inc[3].get("accountId")) != str(out[3].get("accountId")):
                same_text = normalise(out[3].get("description")) == normalise(inc[3].get("description"))
                candidates.append((not _names_other(out[3], inc[3]), not same_text, gap,
                                   out[0], inc[0], out[3], inc[3]))
    candidates.sort(key=lambda c: c[:5])

    found: dict[str, Transfer] = {}
    for *_rank, out_id, in_id, out_txn, in_txn in candidates:
        if out_id in found or in_id in found:
            continue
        owners = frozenset(o for o in (owner_slug(out_txn.get("owner", "")), owner_slug(in_txn.get("owner", ""))) if o)
        found[out_id] = Transfer(out_id, PAIR, owners)
        found[in_id] = Transfer(in_id, PAIR, owners)

    owner_by_iban = {_iban(t.get("ibanMasked")): owner_slug(t.get("owner", ""))
                     for t in transactions if _iban(t.get("ibanMasked"))}
    for tid, _cents_value, _day_value, t in rows:
        if tid in found:
            continue
        other = _iban(t.get("counterpartyAccount"))
        if other and other in owner_by_iban and other != _iban(t.get("ibanMasked")):
            owners = frozenset(o for o in (owner_slug(t.get("owner", "")), owner_by_iban[other]) if o)
            found[tid] = Transfer(tid, OWN_ACCOUNT, owners)
    return found


def transfer_sub_category(transfer: Transfer, joint_owner: str, sub_categories: set[str]) -> str:
    """`From <Person>` when exactly one person (not the joint owner) is involved and the
    taxonomy has that sub-category, else `Internal Transfer`."""
    people = sorted(o for o in transfer.owners if o != joint_owner)
    if len(people) == 1:
        name = "From " + "-".join(part.capitalize() for part in people[0].split("-"))
        if name in sub_categories:
            return name
    return "Internal Transfer"
