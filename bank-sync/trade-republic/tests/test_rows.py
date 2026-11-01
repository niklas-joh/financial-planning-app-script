from zoneinfo import ZoneInfo

from builders import event, with_isin
from tr_export.csv_format import COLUMNS, number, render
from tr_export.rows import classify, to_rows

TZ = ZoneInfo("Europe/Brussels")


def rows(*events):
    return to_rows(events, "Trade Republic", TZ)


def test_classification_by_event_type():
    r = rows(
        event("CARD_TRANSACTION", -4.2, "Demo Bakery"),
        event("CARD_REFUND", 12.99, "Demo Shop"),
        event("CARD_SUCCESSFUL_ATM_WITHDRAWAL", -50, "ATM"),
        event("BANK_TRANSACTION_INCOMING", 500, "Demo Person", "Einzahlung"),
        event("PAYMENT_OUTBOUND", -20, "Demo Utility"),
        with_isin(event("SAVINGS_PLAN_EXECUTED", -100, "Demo World ETF", "Sparplan ausgeführt"), "XS0000000052"),
        event("INTEREST_PAYOUT", 1.5, "Zinsen"),
        event("SOMETHING_NEW", -1, "New"),
    )
    assert [x["Classification"] for x in r.rows] == [
        "card", "card", "card", "transfer", "transfer", "investment", "investment", ""]
    assert r.by_classification == {"card": 3, "transfer": 2, "investment": 2, "unknown": 1}
    assert dict(r.unknown_types) == {"SOMETHING_NEW": 1}
    buy = r.rows[5]
    assert (buy["Type"], buy["ISIN"], buy["Amount"], buy["Subtitle"]) == ("BUY", "XS0000000052", "-100", "Sparplan ausgeführt")


def test_card_note_on_a_migrated_event_is_a_card():
    class Parsed:
        note = "card_successful_transaction"
        event_type = None

    assert classify("TIMELINE_LEGACY_MIGRATED_EVENTS", Parsed()) == "card"


def test_events_without_money_or_cancelled_are_left_out():
    r = rows(
        event("ORDER_CREATED", -10),  # pytr knows it carries no money
        event("CARD_TRANSACTION", None),  # no amount
        event("CARD_TRANSACTION", -3, status="CANCELED"),
        event("ORDER_EXECUTED", -10, details={"sections": [{"type": "header", "data": {"status": "canceled"}}]}),
        event("CARD_TRANSACTION", -2),
    )
    assert len(r.rows) == 1
    assert dict(r.skipped) == {"no money (pytr ignores the type)": 1, "no amount": 1, "cancelled": 2}


def test_booking_date_is_the_brussels_day_and_rows_are_oldest_first():
    late = event("CARD_TRANSACTION", -1, timestamp="2026-08-20T22:30:00.000+0000")  # 00:30 on the 21st in Brussels
    early = event("CARD_TRANSACTION", -1, timestamp="2026-08-19T08:00:00.000+0000")
    r = rows(late, early)
    assert [x["Date"] for x in r.rows] == ["2026-08-19", "2026-08-21"]
    assert r.rows[1]["Timestamp"] == "2026-08-20T22:30:00.000+0000"


def test_saveback_moves_no_cash():
    r = rows(event("BENEFITS_SAVEBACK_EXECUTION", -15, "Demo World ETF", "Saveback"))
    assert (r.rows[0]["Amount"], r.rows[0]["Classification"]) == ("0", "investment")


def test_numbers_are_plain_decimals():
    assert [number(v) for v in (None, 100, -4.2, 0.00001, 1234567.891, -0.0)] == ["", "100", "-4.2", "0.00001", "1234567.891", "0"]


def test_render_writes_the_header_and_semicolon_rows():
    r = rows(event("CARD_TRANSACTION", -4.2, "Bakery; Demo"))
    lines = render(r.rows).decode("utf-8").split("\n")
    assert lines[0] == ";".join(COLUMNS)
    assert '"Bakery; Demo"' in lines[1]
    assert lines[-1] == ""
