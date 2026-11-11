"""Synthetic Trade Republic timeline events (made-up merchants, people, ISINs and ids), shaped like
the events pytr stores in all_events.json."""

from __future__ import annotations

from typing import Any

_n = 0


def event(event_type: str, value: float | None, title: str = "Demo", subtitle: str | None = None,
          timestamp: str = "2026-08-20T10:00:00.000+0000", currency: str = "EUR", **extra: Any) -> dict[str, Any]:
    global _n
    _n += 1
    e: dict[str, Any] = {
        "id": f"0b1e0000-0000-4000-8000-{_n:012d}",
        "timestamp": timestamp,
        "title": title,
        "subtitle": subtitle,
        "eventType": event_type,
        "status": "EXECUTED",
        "source": "timelineTransaction",
    }
    if value is not None:
        e["amount"] = {"currency": currency, "value": value, "fractionDigits": 2}
    e.update(extra)
    return e


def with_isin(e: dict[str, Any], isin: str) -> dict[str, Any]:
    e["icon"] = f"logos/{isin}/v2"
    return e
