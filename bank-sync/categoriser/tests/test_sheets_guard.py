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


@pytest.mark.parametrize("title", ["account_owners", "Financial Planning", "All Transactions"])
def test_refuses_the_owners_tab_and_legacy_names(title):
    with pytest.raises(PermissionError):
        _assert_writable(SimpleNamespace(title=title), SETTINGS)


@pytest.mark.parametrize("title", ["categorised", "categoriser_issues", "categoriser_rules", "categoriser_taxonomy"])
def test_allows_the_job_sheets(title):
    _assert_writable(SimpleNamespace(title=title), SETTINGS)


@pytest.mark.parametrize("field,name", [
    ("output_sheet", "transactions"),
    ("output_sheet", " transactions "),
    ("issues_sheet", "transactions"),
    ("issues_sheet", "categorised"),
    ("rules_sheet", "transactions"),
    ("taxonomy_sheet", "categoriser_rules"),
])
def test_writable_sheets_cannot_collide(field, name):
    with pytest.raises(ValueError):
        Settings(spreadsheet_id="x", credentials_path=Path("/dev/null"), **{field: name})


@pytest.mark.parametrize("value", [0, -0.1, 1.5])
def test_model_min_confidence_must_be_a_probability(value):
    with pytest.raises(ValueError):
        Settings(spreadsheet_id="x", credentials_path=Path("/dev/null"), model_min_confidence=value)
