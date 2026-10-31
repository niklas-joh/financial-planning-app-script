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


def test_allows_the_output_sheet():
    _assert_writable(SimpleNamespace(title="categorised"), SETTINGS)


@pytest.mark.parametrize("name", ["transactions", " transactions "])
def test_output_sheet_cannot_be_the_transactions_sheet(name):
    with pytest.raises(ValueError):
        Settings(spreadsheet_id="x", credentials_path=Path("/dev/null"), output_sheet=name)
