"""`python -m tr_export run` (scheduled, never asks anything) or `python -m tr_export login` (interactive).

Exit codes: 0 done, 3 the Trade Republic session needs a login (run login.sh), 1 any other failure.
The log holds counts, dates and event type names only: never amounts, merchants or account numbers.
"""

from __future__ import annotations

import argparse
import logging
import sys
from typing import Any, Callable

from . import config, drive, fetch, session
from .csv_format import render
from .rows import to_rows

LOGIN_NEEDED = 3
log = logging.getLogger("tr_export")


def run(
    settings: config.Settings,
    open_session: Callable[[config.Settings], Any] = session.resume,
    read_events: Callable[[Any, config.Settings], list] = fetch.events,
    drive_session: Callable[[Any], Any] = drive.session,
) -> int:
    """One export: resume the session, read the timeline, write the CSV, update the Drive file."""
    try:
        tr = open_session(settings)
    except session.LoginNeeded:
        log.error("Trade Republic session missing or expired: run login.sh")
        return LOGIN_NEEDED
    events = read_events(tr, settings)
    result = to_rows(events, settings.account, settings.tz)
    content = render(result.rows)
    settings.data_dir.mkdir(parents=True, exist_ok=True)
    tmp = settings.output_file.with_suffix(".tmp")
    tmp.write_bytes(content)
    tmp.replace(settings.output_file)
    dates = [r["Date"] for r in result.rows]
    log.info(
        "events %d, rows %d (%s), %s; skipped %s",
        len(events),
        len(result.rows),
        ", ".join(f"{k} {v}" for k, v in sorted(result.by_classification.items())) or "none",
        f"{min(dates)} to {max(dates)}" if dates else "no dates",
        ", ".join(f"{k} {v}" for k, v in sorted(result.skipped.items())) or "none",
    )
    if result.unknown_types:
        log.warning("types pytr does not know (rows go to transactions): %s",
                    ", ".join(f"{k} {v}" for k, v in sorted(result.unknown_types.items())))
    if not settings.drive_file_id:
        log.info("TR_DRIVE_FILE_ID not set: wrote %s only (upload it to the CSV folder once, then set the id)",
                 settings.output_file.name)
        return 0
    status = drive.upload(drive_session(settings.credentials), settings.drive_file_id, content)
    log.info("Drive file %s", status)
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="tr_export")
    parser.add_argument("command", choices=("run", "login"))
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    # pytr logs event titles (merchants) at INFO: keep it to warnings.
    logging.getLogger("pytr").setLevel(logging.WARNING)
    try:
        settings = config.load()
        if args.command == "login":
            session.login(settings)
            return 0
        return run(settings)
    except (ValueError, drive.DriveError) as e:
        log.error("%s", e)
        return 1
    except Exception as e:  # noqa: BLE001 - name the failure, never its data
        log.error("failed: %s", type(e).__name__)
        return 1


if __name__ == "__main__":
    sys.exit(main())
