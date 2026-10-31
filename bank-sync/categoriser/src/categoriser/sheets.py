"""Google Sheets I/O.

Hard rule: this module reads `transactions` but only ever writes to the sheets in
`Settings.writable_sheets` (`categorised`, `categoriser_issues`, `categoriser_rules`,
`categoriser_taxonomy`). `_assert_writable` enforces that at every write, and
`Settings` refuses sheet names that collide with the transactions sheet. The
legacy sheet is only ever opened with a read-only scope (`open_readonly`).
"""

from __future__ import annotations

import gspread
from google.oauth2.service_account import Credentials
from gspread.utils import Dimension, ValueInputOption, ValueRenderOption, rowcol_to_a1

from .config import Settings
from .engine import RuleStats
from .issues import ISSUE_COLUMNS, Issue, IssuesPlan, plan_issues
from .merge import UpsertPlan, drop_newly_reviewed, plan_upsert
from .rules import DIRECTIONS, RCOL, RULE_COLUMNS, SOURCES
from .schema import CATEGORISED_COLUMNS, KEY_COLUMN, REVIEWED_COLUMN, CategorisedRow
from .taxonomy import TAXONOMY_COLUMNS

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
READONLY_SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]


def open_spreadsheet(settings: Settings) -> gspread.Spreadsheet:
    creds = Credentials.from_service_account_file(str(settings.credentials_path), scopes=SCOPES)
    return gspread.authorize(creds).open_by_key(settings.spreadsheet_id)


def open_readonly(settings: Settings, spreadsheet_id: str) -> gspread.Spreadsheet:
    """Opens a spreadsheet with a read-only token: Google rejects any write through it."""
    creds = Credentials.from_service_account_file(str(settings.credentials_path), scopes=READONLY_SCOPES)
    return gspread.authorize(creds).open_by_key(spreadsheet_id)


def read_values(ws: gspread.Worksheet) -> list[list]:
    return ws.get_all_values(value_render_option=ValueRenderOption.unformatted)


def _assert_writable(ws: gspread.Worksheet, settings: Settings) -> None:
    if ws.title not in settings.writable_sheets:
        raise PermissionError(f"Refusing to write to sheet '{ws.title}'.")


def read_transactions_header(ss: gspread.Spreadsheet, settings: Settings) -> list[str]:
    return ss.worksheet(settings.transactions_sheet).row_values(1)


def read_transactions(ss: gspread.Spreadsheet, settings: Settings) -> list[dict]:
    ws = ss.worksheet(settings.transactions_sheet)
    return ws.get_all_records(value_render_option=ValueRenderOption.unformatted)


CHECKBOX = {"condition": {"type": "BOOLEAN"}, "strict": True}


def _one_of_list(values: tuple[str, ...]) -> dict:
    return {"condition": {"type": "ONE_OF_LIST", "values": [{"userEnteredValue": v} for v in values]},
            "showCustomUi": True, "strict": True}


def _one_of_range(sheet: str, column: str) -> dict:
    """Dropdown fed by a column of another sheet; not strict, the job validates instead."""
    return {"condition": {"type": "ONE_OF_RANGE",
                          "values": [{"userEnteredValue": f"='{sheet}'!{column}2:{column}"}]},
            "showCustomUi": True, "strict": False}


def _validations(settings: Settings, title: str) -> dict[int, dict]:
    """Checkboxes and dropdowns per column of each job sheet (re-applied after appends,
    since rows added at the bottom don't inherit them)."""
    if title == settings.output_sheet:
        return {REVIEWED_COLUMN: CHECKBOX}
    if title == settings.rules_sheet:
        tax, owners = settings.taxonomy_sheet, settings.owners_sheet
        return {
            RCOL["active"]: CHECKBOX,
            RCOL["direction"]: _one_of_list(DIRECTIONS),
            RCOL["source"]: _one_of_list(SOURCES),
            RCOL["owner"]: _one_of_range(owners, "A"),
            RCOL["type"]: _one_of_range(tax, "A"),
            RCOL["category"]: _one_of_range(tax, "B"),
            RCOL["subCategory"]: _one_of_range(tax, "C"),
        }
    return {}


def ensure_categorised_sheet(ss: gspread.Spreadsheet, settings: Settings) -> gspread.Worksheet:
    """Return the output sheet, creating it with headers and a checkbox column if missing."""
    return _ensure_sheet(ss, settings, settings.output_sheet, CATEGORISED_COLUMNS)


def ensure_issues_sheet(ss: gspread.Spreadsheet, settings: Settings) -> gspread.Worksheet:
    """Return the issues log sheet, creating it with headers if missing."""
    return _ensure_sheet(ss, settings, settings.issues_sheet, ISSUE_COLUMNS)


def ensure_taxonomy_sheet(ss: gspread.Spreadsheet, settings: Settings) -> gspread.Worksheet:
    """Return the taxonomy sheet, creating it (header only) if missing. Seed it with `seed_taxonomy`."""
    return _ensure_sheet(ss, settings, settings.taxonomy_sheet, TAXONOMY_COLUMNS)


def ensure_rules_sheet(ss: gspread.Spreadsheet, settings: Settings) -> gspread.Worksheet:
    """Return the rules sheet, creating it with a checkbox on `active` and dropdowns on
    `direction`, `source`, `owner` (from the account owners tab) and the three
    category columns (from the taxonomy sheet)."""
    return _ensure_sheet(ss, settings, settings.rules_sheet, RULE_COLUMNS)


def _ensure_sheet(ss: gspread.Spreadsheet, settings: Settings, title: str,
                  columns: tuple[str, ...]) -> gspread.Worksheet:
    try:
        ws = ss.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = ss.add_worksheet(title=title, rows=1000, cols=len(columns))
        _init_sheet(ss, ws, settings, columns)
        return ws

    header = ws.row_values(1)
    if not header:
        _init_sheet(ss, ws, settings, columns)
    elif header[: len(columns)] != list(columns):
        raise ValueError(f"Sheet '{title}' exists with unexpected headers: {header}")
    return ws


def _init_sheet(ss: gspread.Spreadsheet, ws: gspread.Worksheet, settings: Settings,
                columns: tuple[str, ...]) -> None:
    """Write the header row, freeze it and add the sheet's checkboxes and dropdowns."""
    _assert_writable(ws, settings)
    ws.update([list(columns)], "A1", value_input_option=ValueInputOption.raw)
    ws.freeze(rows=1)
    _apply_validations(ss, ws, settings)


def _apply_validations(ss: gspread.Spreadsheet, ws: gspread.Worksheet, settings: Settings) -> None:
    validations = _validations(settings, ws.title)
    if not validations:
        return
    ss.batch_update({"requests": [{
        "setDataValidation": {
            "range": {
                "sheetId": ws.id,
                "startRowIndex": 1,
                "startColumnIndex": column,
                "endColumnIndex": column + 1,
            },
            "rule": rule,
        }
    } for column, rule in sorted(validations.items())]})


def append_rows(ws: gspread.Worksheet, settings: Settings, rows: list[list]) -> None:
    """Append rows below the existing ones (mined rules, taxonomy seed)."""
    _assert_writable(ws, settings)
    _write_rows(ws, settings, {}, rows, 0)


def write_rule_stats(ws: gspread.Worksheet, settings: Settings, stats: dict[str, RuleStats]) -> int:
    """Fill `hits` and `lastMatched` for each rule; touches only those two columns.

    The sheet is read just before writing and rows are found by ruleId, so rules
    added, moved or deleted from the phone meanwhile are not overwritten. A ruleId
    on more than one row is skipped. Returns the number of rows changed.
    """
    _assert_writable(ws, settings)
    values = read_values(ws)
    rows: dict[str, list[int]] = {}
    for n, row in enumerate(values[1:], start=2):
        rid = str(row[0]).strip() if row else ""
        if rid:
            rows.setdefault(rid, []).append(n)
    hits_col, last_col = RCOL["hits"], RCOL["lastMatched"]
    data = []
    for rid, st in stats.items():
        if len(rows.get(rid, [])) != 1:
            continue
        n = rows[rid][0]
        row = values[n - 1]
        current = [str(row[c]) if len(row) > c else "" for c in (hits_col, last_col)]
        if current != [str(st.hits), st.lastMatched]:
            data.append({"range": f"{rowcol_to_a1(n, hits_col + 1)}:{rowcol_to_a1(n, last_col + 1)}",
                         "values": [[st.hits, st.lastMatched]]})
    if data:
        ws.batch_update(data, value_input_option=ValueInputOption.raw)
    return len(data)


def _write_rows(ws: gspread.Worksheet, settings: Settings, updates: dict[int, list], appends: list[list],
                width: int) -> None:
    """One batch_update for changed rows; new rows go right below the last row with a key.

    Not `append_rows`: an empty checkbox reads as FALSE, so Sheets' own append would
    treat a sheet of blank checkbox rows as full and write below row 1000.
    """
    if updates:
        ws.batch_update(
            [{"range": f"{rowcol_to_a1(r, 1)}:{rowcol_to_a1(r, width)}", "values": [v]}
             for r, v in sorted(updates.items())],
            value_input_option=ValueInputOption.raw,
        )
    if appends:
        keys = ws.col_values(1, value_render_option=ValueRenderOption.unformatted)
        start = max((n for n, v in enumerate(keys, start=1) if str(v).strip()), default=1) + 1
        end = start + len(appends) - 1
        if end > ws.row_count:
            ws.add_rows(end - ws.row_count)
            _apply_validations(ws.spreadsheet, ws, settings)
        ws.update(appends, rowcol_to_a1(start, 1), value_input_option=ValueInputOption.raw)


def record_issues(ws: gspread.Worksheet, settings: Settings, current: list[Issue],
                  checked_types: set[str], now: str) -> IssuesPlan:
    """Log new issues, refresh ones still present, resolve ones that are gone."""
    _assert_writable(ws, settings)
    existing = ws.get_all_values(value_render_option=ValueRenderOption.unformatted)
    plan = plan_issues(existing, current, checked_types, now)
    _write_rows(ws, settings, plan.updates, plan.appends, len(ISSUE_COLUMNS))
    return plan


def _column_letter(index: int) -> str:
    return rowcol_to_a1(1, index + 1).rstrip("1")


def upsert_categorised(ws: gspread.Worksheet, settings: Settings, rows: list[CategorisedRow],
                       clear_stale_at: str | None = None) -> UpsertPlan:
    """Write rows keyed by transactionId; rows marked reviewed are left untouched.

    The key and reviewed columns are re-read just before writing, so a row
    ticked (or moved) while the run was planning is skipped, not overwritten.
    A tick made during the write call itself can still be lost; it is a
    single request lasting well under a second.
    """
    _assert_writable(ws, settings)
    existing = ws.get_all_values(value_render_option=ValueRenderOption.unformatted)
    plan = plan_upsert(existing, rows, clear_stale_at)

    if plan.updates:
        key_col, reviewed_col = _column_letter(KEY_COLUMN), _column_letter(REVIEWED_COLUMN)
        keys, reviewed = ws.batch_get(
            [f"{key_col}:{key_col}", f"{reviewed_col}:{reviewed_col}"],
            major_dimension=Dimension.cols,
            value_render_option=ValueRenderOption.unformatted,
        )
        plan = drop_newly_reviewed(plan, keys[0] if keys else [], reviewed[0] if reviewed else [])
    _write_rows(ws, settings, plan.updates, plan.appends, len(CATEGORISED_COLUMNS))
    return plan
