"""Reads the timeline with pytr and keeps every event in `data/all_events.json` (pytr's event database).

The first run reads the whole history; later runs read only the last `TR_LOOKBACK_DAYS` days and pytr
merges them into the database (newer copies of an event replace older ones), so a run stays short and
the CSV still covers everything.
"""

from __future__ import annotations

import asyncio
from datetime import datetime, timedelta, timezone
from typing import Any

from pytr.timeline import Timeline

from .config import Settings

DATABASE = "all_events.json"


def since(settings: Settings, now: datetime) -> float:
    """Unix time to read from: 0 (everything) without a database, else `now - lookback`."""
    if not (settings.data_dir / DATABASE).exists():
        return 0.0
    return (now - timedelta(days=settings.lookback_days)).timestamp()


def events(tr: Any, settings: Settings, now: datetime | None = None) -> list[dict[str, Any]]:
    """Every stored event after this run's read (the database is updated on disk)."""
    settings.data_dir.mkdir(parents=True, exist_ok=True)
    timeline = Timeline(
        tr,
        settings.data_dir,
        not_before=since(settings, now or datetime.now(timezone.utc)),
        store_event_database=True,
    )
    asyncio.run(timeline.tl_loop())
    return list(timeline.events)
