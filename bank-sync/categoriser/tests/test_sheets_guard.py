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


def test_output_sheet_cannot_be_pointed_at_transactions():
    bad = Settings(spreadsheet_id="x", credentials_path=Path("/dev/null"), output_sheet="transactions")
    with pytest.raises(PermissionError):
        _assert_writable(SimpleNamespace(title="transactions"), bad)
