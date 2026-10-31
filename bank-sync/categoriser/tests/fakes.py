"""In-memory stand-ins for the parts of gspread the categoriser uses."""

from __future__ import annotations

import re

import gspread
from gspread.utils import a1_to_rowcol


class FakeWorksheet:
    def __init__(self, title: str, values: list[list] | None = None, sheet_id: int = 1):
        self.title = title
        self.id = sheet_id
        self.values = [list(r) for r in (values or [])]
        self.frozen_rows = 0
        self.before_batch_get = None  # hook: simulate an edit between plan and write
        self.calls: list[str] = []

    # reads
    def get_all_values(self, **_):
        self.calls.append("get_all_values")
        return [list(r) for r in self.values]

    def row_values(self, row, **_):
        return list(self.values[row - 1]) if len(self.values) >= row else []

    def batch_get(self, ranges, **_):
        self.calls.append("batch_get")
        if self.before_batch_get:
            self.before_batch_get(self)
        out = []
        for rng in ranges:
            col = re.match(r"([A-Z]+):\1$", rng).group(1)
            idx = a1_to_rowcol(f"{col}1")[1] - 1
            out.append([[r[idx] if len(r) > idx else "" for r in self.values]])
        return out

    # writes
    def update(self, values, range_name, **_):
        self.calls.append("update")
        row, col = a1_to_rowcol(range_name)
        for i, vals in enumerate(values):
            self._set_row(row + i, col, vals)

    def batch_update(self, data, **_):
        self.calls.append("batch_update")
        for item in data:
            row, col = a1_to_rowcol(item["range"].split(":")[0])
            self._set_row(row, col, item["values"][0])

    def append_rows(self, rows, **_):
        self.calls.append("append_rows")
        self.values.extend(list(r) for r in rows)

    def freeze(self, rows=0, **_):
        self.frozen_rows = rows

    def _set_row(self, row, col, vals):
        while len(self.values) < row:
            self.values.append([])
        target = self.values[row - 1]
        while len(target) < col - 1 + len(vals):
            target.append("")
        target[col - 1: col - 1 + len(vals)] = list(vals)


class FakeSpreadsheet:
    def __init__(self, sheets: dict[str, FakeWorksheet] | None = None):
        self.sheets = dict(sheets or {})
        self.requests: list[dict] = []

    def worksheet(self, title):
        if title not in self.sheets:
            raise gspread.WorksheetNotFound(title)
        return self.sheets[title]

    def add_worksheet(self, title, rows, cols):
        ws = FakeWorksheet(title, sheet_id=len(self.sheets) + 1)
        self.sheets[title] = ws
        return ws

    def batch_update(self, body):
        self.requests.extend(body["requests"])
