import logging

import pytest

from builders import event
from fakes import FakeDrive
from tr_export import __main__ as main
from tr_export import config, fetch, session


def settings(tmp_path, **env):
    return config.load({"TR_PHONE": "+352000000000", "TR_HOME": str(tmp_path), **env})


EVENTS = [event("CARD_TRANSACTION", -4.2, "Demo Bakery"), event("SOMETHING_NEW", -1)]


def test_run_writes_the_csv_and_logs_counts_never_merchants_or_amounts(tmp_path, caplog):
    s = settings(tmp_path)
    caplog.set_level(logging.INFO)
    code = main.run(s, open_session=lambda _: object(), read_events=lambda tr, _: EVENTS,
                    drive_session=lambda _: pytest.fail("no upload without TR_DRIVE_FILE_ID"))
    assert code == 0
    text = s.output_file.read_text(encoding="utf-8")
    assert text.startswith("TR event id;") and "Demo Bakery" in text
    log = caplog.text
    assert "rows 2 (card 1, unknown 1)" in log and "SOMETHING_NEW 1" in log
    assert "Demo Bakery" not in log and "4.2" not in log


def test_run_uploads_to_the_drive_file(tmp_path):
    s = settings(tmp_path, TR_DRIVE_FILE_ID="file-1")
    http = FakeDrive(current=b"")
    run = lambda: main.run(s, open_session=lambda _: object(), read_events=lambda tr, _: EVENTS, drive_session=lambda _: http)
    assert run() == 0 and run() == 0
    assert len(http.patches) == 1 and http.current == s.output_file.read_bytes()


def test_expired_session_exits_3_and_writes_nothing(tmp_path):
    s = settings(tmp_path)

    def expired(_):
        raise session.LoginNeeded("expired")

    assert main.run(s, open_session=expired, read_events=lambda *a: pytest.fail("no read")) == main.LOGIN_NEEDED
    assert not s.output_file.exists()


def test_first_run_reads_everything_later_runs_the_lookback(tmp_path):
    from datetime import datetime, timezone
    s = settings(tmp_path, TR_LOOKBACK_DAYS="10")
    now = datetime(2026, 9, 1, tzinfo=timezone.utc)
    assert fetch.since(s, now) == 0.0
    s.data_dir.mkdir(parents=True)
    (s.data_dir / fetch.DATABASE).write_text("[]")
    assert fetch.since(s, now) == datetime(2026, 8, 22, tzinfo=timezone.utc).timestamp()


def test_main_reports_a_missing_setting_by_name(monkeypatch, caplog):
    monkeypatch.delenv("TR_PHONE", raising=False)
    assert main.main(["run"]) == 1
    assert "TR_PHONE" in caplog.text


def test_events_reads_the_timeline_into_the_event_database(tmp_path, monkeypatch):
    s = settings(tmp_path)
    seen = {}

    class FakeTimeline:
        def __init__(self, tr, output_path, not_before, store_event_database):
            seen.update(tr=tr, path=output_path, not_before=not_before, store=store_event_database)
            self.events = EVENTS

        async def tl_loop(self):
            seen["looped"] = True

    monkeypatch.setattr(fetch, "Timeline", FakeTimeline)
    tr = object()
    assert fetch.events(tr, s) == EVENTS
    assert seen == {"tr": tr, "path": s.data_dir, "not_before": 0.0, "store": True, "looped": True}
