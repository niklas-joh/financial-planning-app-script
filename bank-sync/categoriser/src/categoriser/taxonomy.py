"""Allowed (type, category, subCategory) combinations, kept in the `categoriser_taxonomy` sheet.

The owner maintains that sheet; the job only creates it and seeds it once from the
legacy sheet's `Dropdowns` tab (first three columns). Rule outputs must be one of
these combinations.
"""

from __future__ import annotations

from dataclasses import dataclass

TAXONOMY_COLUMNS: tuple[str, ...] = ("type", "category", "subCategory")

Triple = tuple[str, str, str]


@dataclass(frozen=True)
class Taxonomy:
    triples: frozenset[Triple]

    def __contains__(self, triple: Triple) -> bool:
        return tuple(triple) in self.triples

    def __len__(self) -> int:
        return len(self.triples)


def _triples(rows: list[list]) -> list[Triple]:
    out: list[Triple] = []
    seen: set[Triple] = set()
    for row in rows:
        cells = [str(c).strip() for c in list(row[:3]) + [""] * (3 - len(row[:3]))]
        triple = (cells[0], cells[1], cells[2])
        if all(triple) and triple not in seen:
            seen.add(triple)
            out.append(triple)
    return out


def parse_taxonomy(values: list[list]) -> Taxonomy:
    """`values` is the whole taxonomy sheet incl. header. Incomplete rows are ignored."""
    if values and [str(h).strip() for h in values[0][:3]] != list(TAXONOMY_COLUMNS):
        raise ValueError("Header row of the taxonomy sheet does not match the expected columns.")
    return Taxonomy(frozenset(_triples(values[1:])))


def triples_from_dropdowns(values: list[list]) -> list[Triple]:
    """Rows of the legacy `Dropdowns` tab (header first); columns A:C are Type, Category, Sub-Category."""
    return sorted(_triples(values[1:]))
