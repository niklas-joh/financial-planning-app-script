import pytest

from categoriser.merge import plan_upsert
from categoriser.schema import CATEGORISED_COLUMNS, CategorisedRow, Method

HEADER = list(CATEGORISED_COLUMNS)


def row(tid, category="Groceries", reviewed=False):
    return CategorisedRow(tid, "joint", "Essentials", category, "", Method.RULE, 1.0, "R1", reviewed,
                          "2026-10-30T06:15:00Z")


def sheet_row(tid, reviewed):
    return [tid, "joint", "Essentials", "Old", "", "rule", 1, "R1", reviewed, "2026-10-01T00:00:00Z"]


def test_new_ids_are_appended():
    plan = plan_upsert([HEADER], [row("a"), row("b")])
    assert [v[0] for v in plan.appends] == ["a", "b"]
    assert plan.updates == {}


def test_unreviewed_rows_are_updated_in_place():
    plan = plan_upsert([HEADER, sheet_row("a", False)], [row("a")])
    assert plan.updates[2][3] == "Groceries"
    assert plan.appends == []


@pytest.mark.parametrize("reviewed", [True, "TRUE", "true"])
def test_reviewed_rows_are_never_overwritten(reviewed):
    plan = plan_upsert([HEADER, sheet_row("a", reviewed)], [row("a")])
    assert plan.updates == {} and plan.appends == []
    assert plan.skipped_reviewed == ["a"]


def test_empty_sheet_without_header_appends():
    assert len(plan_upsert([], [row("a")]).appends) == 1


def test_wrong_header_is_rejected():
    with pytest.raises(ValueError):
        plan_upsert([["id", "category"]], [row("a")])


def test_duplicate_ids_in_one_run_are_rejected():
    with pytest.raises(ValueError):
        plan_upsert([HEADER], [row("a"), row("a")])


def test_method_is_written_as_plain_string():
    values = row("a").to_values()
    assert values[CATEGORISED_COLUMNS.index("method")] == "rule"
    assert type(values[CATEGORISED_COLUMNS.index("method")]) is str
