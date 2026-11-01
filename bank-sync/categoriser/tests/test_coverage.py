from categoriser.coverage import NOT_UNIQUE, NOT_WRITTEN, ORPHANED, check_coverage
from categoriser.schema import CATEGORISED_COLUMNS

HEADER = list(CATEGORISED_COLUMNS)


def row(tid, label=("Essentials", "Groceries", "Supermarket"), reviewed=False):
    return [tid, "alice", *label, "rule", "", "R1", reviewed, "2026-10-01T00:00:00Z"]


def blank_checkbox_row():
    """What a never-used row of `categorised` reads as: an unticked checkbox and nothing else."""
    return ["", "", "", "", "", "", "", "", False, ""]


def check(transactions, categorised, produced=()):
    return check_coverage([{"transactionId": t} for t in transactions], [HEADER, *categorised], set(produced),
                          "transactions", "categorised")


def test_counts_categorised_and_uncategorised_transactions():
    cov = check(["a", "b", "c"], [row("a"), row("b", label=("", "", ""))])
    assert (cov.transactions, cov.categorised, cov.uncategorised) == (3, 1, 2)
    assert cov.issues == []


def test_blank_checkbox_rows_are_not_orphans():
    cov = check(["a"], [row("a")] + [blank_checkbox_row()] * 306)
    assert cov.orphaned == 0 and cov.issues == []


def test_row_without_a_transaction_is_reported_as_orphaned_and_says_if_reviewed():
    cov = check(["a"], [row("a"), row("gone", reviewed=True)])
    assert cov.orphaned == 1
    issue = cov.issues[0]
    assert (issue.type, issue.transactionId) == (ORPHANED, "gone")
    assert "Row 3" in issue.detail and "reviewed" in issue.detail


def test_labelled_row_missing_after_the_write_is_reported():
    cov = check(["a", "b"], [row("a")], produced={"a", "b"})
    assert cov.not_written == 1
    assert [(i.type, i.transactionId) for i in cov.issues] == [(NOT_WRITTEN, "b")]


def test_transaction_id_on_several_transactions_rows_is_reported_once():
    cov = check(["a", "a", "b"], [row("a")])
    assert cov.transactions == 2 and cov.not_unique == 1
    assert [(i.type, i.transactionId) for i in cov.issues] == [(NOT_UNIQUE, "a")]
    assert "2 rows" in cov.issues[0].detail
