import pytest

from categoriser.taxonomy import TAXONOMY_COLUMNS, parse_taxonomy, triples_from_dropdowns


def test_parses_complete_rows_and_ignores_the_rest():
    tax = parse_taxonomy([list(TAXONOMY_COLUMNS), ["A", "B", "C"], ["A", "B", ""], [" A ", "B", "C"], ["X", "Y", "Z", "extra"]])
    assert tax.triples == {("A", "B", "C"), ("X", "Y", "Z")}
    assert ("A", "B", "C") in tax and ("A", "B", "") not in tax


def test_rejects_a_renamed_header():
    with pytest.raises(ValueError):
        parse_taxonomy([["type", "cat", "subCategory"]])


def test_empty_sheet_is_an_empty_taxonomy():
    assert len(parse_taxonomy([])) == 0


def test_dropdowns_tab_uses_the_first_three_columns():
    values = [["Type", "Category", "Sub-Category", "", "Unique"], ["B", "x", "y", "", "junk"], ["A", "x", "y"], ["", "", ""]]
    assert triples_from_dropdowns(values) == [("A", "x", "y"), ("B", "x", "y")]
