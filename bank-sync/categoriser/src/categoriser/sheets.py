"""Google Sheets I/O.

Hard rule: this module reads `transactions` but only ever writes to the output
sheet (`categorised`). `_assert_writable` enforces that at every write, and
`Settings` refuses an output sheet named like the transactions sheet.
"""

from __future__ import annotations

import gspread
from google.oauth2.service_account import Credentials
from gspread.utils import Dimension, ValueInputOption, ValueRenderOption, rowcol_to_a1

from .config import Settings
from .merge import UpsertPlan, drop_newly_reviewed, plan_upsert
from .schema import CATEGORISED_COLUMNS, KEY_COLUMN, REVIEWED_COLUMN, CategorisedRow

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def open_spreadsheet(settings: Settings) -> gspread.Spreadsheet:
    creds = Credentials.from_service_account_file(str(settings.credentials_path), scopes=SCOPES)
    return gspread.authorize(creds).open_by_key(settings.spreadsheet_id)


def _assert_writable(ws: gspread.Worksheet, settings: Settings) -> None:
    if ws.title != settings.output_sheet:
        raise PermissionError(f"Refusing to write to sheet '{ws.title}'.")


def read_transactions_header(ss: gspread.Spreadsheet, settings: Settings) -> list[str]:
    return ss.worksheet(settings.transactions_sheet).row_values(1)


def read_transactions(ss: gspread.Spreadsheet, settings: Settings) -> list[dict]:
    ws = ss.worksheet(settings.transactions_sheet)
    return ws.get_all_records(value_render_option=ValueRenderOption.unformatted)


def ensure_categorised_sheet(ss: gspread.Spreadsheet, settings: Settings) -> gspread.Worksheet:
    """Return the output sheet, creating it with headers and a checkbox column if missing."""
    try:
        ws = ss.worksheet(settings.output_sheet)
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(title=settings.output_sheet, rows=1000, cols=len(CATEGORISED_COLUMNS))
        _init_sheet(ss, ws, settings)
        return ws

    header = ws.row_values(1)
    if not header:
        _init_sheet(ss, ws, settings)
    elif header[: len(CATEGORISED_COLUMNS)] != list(CATEGORISED_COLUMNS):
        raise ValueError(f"Sheet '{settings.output_sheet}' exists with unexpected headers: {header}")
    return ws


def _init_sheet(ss: gspread.Spreadsheet, ws: gspread.Worksheet, settings: Settings) -> None:
    """Write the header row, freeze it and make `reviewed` a checkbox column."""
    _assert_writable(ws, settings)
    ws.update([list(CATEGORISED_COLUMNS)], "A1", value_input_option=ValueInputOption.raw)
    ws.freeze(rows=1)
    ss.batch_update({"requests": [{
        "setDataValidation": {
            "range": {
                "sheetId": ws.id,
                "startRowIndex": 1,
                "startColumnIndex": REVIEWED_COLUMN,
                "endColumnIndex": REVIEWED_COLUMN + 1,
            },
            "rule": {"condition": {"type": "BOOLEAN"}, "strict": True},
        }
    }]})


def _column_letter(index: int) -> str:
    return rowcol_to_a1(1, index + 1).rstrip("1")


def upsert_categorised(ws: gspread.Worksheet, settings: Settings, rows: list[CategorisedRow]) -> UpsertPlan:
    """Write rows keyed by transactionId; rows marked reviewed are left untouched.

    The key and reviewed columns are re-read just before writing, so a row
    ticked (or moved) while the run was planning is skipped, not overwritten.
    A tick made during the write call itself can still be lost; it is a
    single request lasting well under a second.
    """
    _assert_writable(ws, settings)
    existing = ws.get_all_values(value_render_option=ValueRenderOption.unformatted)
    plan = plan_upsert(existing, rows)

    last_col = len(CATEGORISED_COLUMNS)
    if plan.updates:
        key_col, reviewed_col = _column_letter(KEY_COLUMN), _column_letter(REVIEWED_COLUMN)
        keys, reviewed = ws.batch_get(
            [f"{key_col}:{key_col}", f"{reviewed_col}:{reviewed_col}"],
            major_dimension=Dimension.cols,
            value_render_option=ValueRenderOption.unformatted,
        )
        plan = drop_newly_reviewed(plan, keys[0] if keys else [], reviewed[0] if reviewed else [])
    if plan.updates:
        ws.batch_update(
            [{"range": f"{rowcol_to_a1(r, 1)}:{rowcol_to_a1(r, last_col)}", "values": [v]}
             for r, v in sorted(plan.updates.items())],
            value_input_option=ValueInputOption.raw,
        )
    if plan.appends:
        ws.append_rows(plan.appends, value_input_option=ValueInputOption.raw, table_range="A1")
    return plan
