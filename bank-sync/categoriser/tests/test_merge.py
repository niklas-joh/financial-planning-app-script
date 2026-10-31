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


def test_duplicate_ids_in_the_sheet_are_skipped_and_reported():
    existing = [HEADER, sheet_row("a", False), sheet_row("b", False), sheet_row("a", False)]
    plan = plan_upsert(existing, [row("a"), row("b")])
    assert plan.duplicates == {"a": [2, 4]}
    assert list(plan.updates) == [3]          # b still updated
    assert plan.appends == []                 # a is not appended a third time


def test_missing_confidence_is_written_as_blank():
    manual = CategorisedRow("a", "me", "Wants/Pleasure", "Dining", "", Method.MANUAL, None, "", True, "t")
    assert manual.to_values()[CATEGORISED_COLUMNS.index("confidence")] == ""


def test_method_is_written_as_plain_string():
    values = row("a").to_values()
    assert values[CATEGORISED_COLUMNS.index("method")] == "rule"
    assert type(values[CATEGORISED_COLUMNS.index("method")]) is str


def test_stale_unreviewed_rule_rows_are_cleared():
    from categoriser.merge import drop_newly_reviewed
    existing = [HEADER, sheet_row("gone", False), sheet_row("kept-reviewed", True),
                ["manual-row", "joint", "Essentials", "X", "", "manual", "", "", False, ""], sheet_row("still", False)]
    plan = plan_upsert(existing, [row("still")], clear_stale_at="2026-10-31T00:00:00Z")
    assert plan.cleared == ["gone"]
    assert plan.updates[2] == ["gone", "", "", "", "", "", "", "", False, "2026-10-31T00:00:00Z"]
    assert set(plan.updates) == {2, 5}
    plan = drop_newly_reviewed(plan, ["transactionId", "gone", "kept-reviewed", "manual-row", "still"],
                               ["reviewed", True, True, False, False])
    assert plan.cleared == [] and plan.skipped_reviewed == ["gone"]


def test_no_clearing_without_a_timestamp():
    assert plan_upsert([HEADER, sheet_row("gone", False)], []).updates == {}


def test_unchanged_rows_are_not_rewritten():
    same = row("a", category="Old")
    existing = [HEADER, ["a", "joint", "Essentials", "Old", "", "rule", 1, "R1", False, "2026-01-01T00:00:00Z"]]
    assert plan_upsert(existing, [same]).updates == {}


def test_stale_model_rows_are_cleared_only_when_asked():
    model_row = ["m", "joint", "Essentials", "Old", "", "model", 0.95, "", False, "2026-10-01T00:00:00Z"]
    existing = [HEADER, model_row, sheet_row("r", False)]
    plan = plan_upsert(existing, [], clear_stale_at="2026-10-31T00:00:00Z")
    assert plan.cleared == ["r"]                                  # model rows kept by default
    plan = plan_upsert(existing, [], clear_stale_at="2026-10-31T00:00:00Z",
                       clear_methods=frozenset({Method.RULE, Method.MODEL}))
    assert sorted(plan.cleared) == ["m", "r"]
