from pathlib import Path
from types import SimpleNamespace

import pytest

from categoriser.config import Settings
from categoriser.sheets import _assert_writable

SETTINGS = Settings(spreadsheet_id="x", credentials_path=Path("/dev/null"))


@pytest.mark.parametrize("title", ["transactions", "Sheet1", "accounts"])
def test_refuses_to_write_anywhere_but_the_output_sheet(title):
    with pytest.raises(PermissionError):
        _assert_writable(SimpleNamespace(title=title), SETTINGS)


@pytest.mark.parametrize("title", ["categorised", "categoriser_issues"])
def test_allows_the_output_and_issues_sheets(title):
    _assert_writable(SimpleNamespace(title=title), SETTINGS)


@pytest.mark.parametrize("field,name", [
    ("output_sheet", "transactions"),
    ("output_sheet", " transactions "),
    ("issues_sheet", "transactions"),
    ("issues_sheet", "categorised"),
])
def test_writable_sheets_cannot_collide(field, name):
    with pytest.raises(ValueError):
        Settings(spreadsheet_id="x", credentials_path=Path("/dev/null"), **{field: name})
