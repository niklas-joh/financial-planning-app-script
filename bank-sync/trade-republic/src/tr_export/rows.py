"""Pure: Trade Republic timeline events (as pytr stores them in `all_events.json`) to CSV rows.

Which tab a row belongs on is decided here, from the event type, and written to `Classification`:

- `card`: card spending, refunds and ATM withdrawals (every `CARD_*` event, and pytr's card notes on
  migrated events) -> Bank Sync's `transactions` tab.
- `transfer`: money in or out of the cash account that is not a card (SEPA transfers, direct debits,
  top-ups) -> `transactions`, so both legs of a bank transfer are there to pair.
- `investment`: trades, savings plans, round-ups, saveback, dividends, interest, taxes -> `investment_transactions`.
- `''`: a type pytr does not know yet -> `transactions`, so it is seen; the run log names the type.

Events without an amount, cancelled events and the types pytr knows carry no money (order created,
documents, reports ...) are left out.
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Iterable, Mapping
from zoneinfo import ZoneInfo

from pytr.event import ConditionalEventType, Event, PPEventType, events_known_ignored

from .csv_format import number

CARD = "card"
TRANSFER = "transfer"
INVESTMENT = "investment"
UNKNOWN = ""

# pytr's notes for card events it recognises by their details (e.g. migrated "Kartenzahlung").
_CARD_NOTES = {"card_successful_transaction", "card_refund"}
_CASH_TYPES = {PPEventType.DEPOSIT, PPEventType.REMOVAL}


@dataclass
class Result:
    """Rows to write and what was left out (counts and type names only: no amounts or texts)."""

    rows: list[dict[str, str]] = field(default_factory=list)
    by_classification: Counter = field(default_factory=Counter)
    skipped: Counter = field(default_factory=Counter)
    unknown_types: Counter = field(default_factory=Counter)


def parse_timestamp(text: str) -> datetime:
    """TR timestamps look like `2026-08-20T21:58:03.120+0000`."""
    try:
        return datetime.strptime(text, "%Y-%m-%dT%H:%M:%S.%f%z")
    except ValueError:
        return datetime.fromisoformat(text)


def _parsed(event: Mapping[str, Any]) -> Event | None:
    """pytr's reading of the event (type, ISIN, shares, fees, taxes), or None if pytr cannot read it."""
    try:
        return Event.from_dict(dict(event))
    except Exception:  # noqa: BLE001 - one odd event must not stop the export; it is still written, unclassified
        return None


def classify(event_type: str, parsed: Event | None) -> str:
    """`card`, `transfer`, `investment` or '' (unknown) for one event."""
    if event_type.startswith("CARD_") or (parsed is not None and parsed.note in _CARD_NOTES):
        return CARD
    kind = parsed.event_type if parsed is not None else None
    if kind is None:
        return UNKNOWN
    if kind in _CASH_TYPES:
        return TRANSFER
    if isinstance(kind, (PPEventType, ConditionalEventType)):
        return INVESTMENT
    return UNKNOWN


def _cancelled(event: Mapping[str, Any]) -> bool:
    """Cancelled on the event or in its detail header (the two places pytr checks)."""
    if str(event.get("status") or "").upper() in ("CANCELED", "CANCELLED"):
        return True
    for section in (event.get("details") or {}).get("sections") or []:
        if isinstance(section, dict) and section.get("type") == "header":
            if str((section.get("data") or {}).get("status") or "").upper() in ("CANCELED", "CANCELLED"):
                return True
    return False


def _skip_reason(event: Mapping[str, Any], event_type: str) -> str | None:
    if _cancelled(event):
        return "cancelled"
    if event_type in events_known_ignored:
        return "no money (pytr ignores the type)"
    amount = event.get("amount") or {}
    if amount.get("value") is None or not amount.get("currency"):
        return "no amount"
    return None


def to_rows(events: Iterable[Mapping[str, Any]], account: str, tz: ZoneInfo) -> Result:
    """Every event with money to one CSV row, oldest first.

    Args:
        events: Raw timeline events (each needs `id`, `timestamp`, `title`, `subtitle`, `eventType`, `amount`).
        account: Account name written to every row (one Trade Republic cash account per person).
        tz: Time zone of the booking date (Bank Sync's clock).
    """
    result = Result()
    for event in sorted(events, key=lambda e: parse_timestamp(e["timestamp"])):
        event_type = str(event.get("eventType") or "").upper()
        reason = _skip_reason(event, event_type)
        if reason:
            result.skipped[reason] += 1
            continue
        parsed = _parsed(event)
        classification = classify(event_type, parsed)
        amount = event["amount"]
        value = amount["value"]
        # Saveback shares are a gift: TR shows their value, but no cash leaves the account.
        if parsed is not None and parsed.event_type == ConditionalEventType.SAVEBACK:
            value = 0
        stamp = parse_timestamp(event["timestamp"])
        result.rows.append({
            "TR event id": str(event["id"]),
            "Timestamp": event["timestamp"],
            "Date": stamp.astimezone(tz).date().isoformat(),
            "Account": account,
            "Event type": event_type,
            "Type": _type_name(parsed),
            "Title": str(event.get("title") or ""),
            "Subtitle": str(event.get("subtitle") or ""),
            "Amount": number(value),
            "Currency": str(amount["currency"]).upper(),
            "ISIN": (parsed.isin if parsed is not None else None) or "",
            "Shares": number(parsed.shares if parsed is not None else None),
            "Fees": number(parsed.fees if parsed is not None else None),
            "Taxes": number(parsed.taxes if parsed is not None else None),
            "Classification": classification,
        })
        result.by_classification[classification or "unknown"] += 1
        if classification == UNKNOWN:
            result.unknown_types[event_type or "(none)"] += 1
    return result


def _type_name(parsed: Event | None) -> str:
    """pytr's kind of movement (BUY, SELL, DIVIDEND, DEPOSIT, ...), or ''."""
    kind = parsed.event_type if parsed is not None else None
    if kind is None:
        return ""
    if kind == ConditionalEventType.TRADE_INVOICE and parsed.value is not None:
        return "BUY" if parsed.value < 0 else "SELL"
    return kind.value if isinstance(kind, PPEventType) else kind.name
