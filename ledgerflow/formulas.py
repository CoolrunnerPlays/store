"""Rewriting of A1 cell references when rows are inserted into a sheet.

openpyxl moves cells when you insert rows but leaves every formula in the file
pointing at the old addresses, so a workbook edited naively comes back with its
totals silently wrong. This module does what Excel itself would do, plus the one
thing Excel does not: a range whose bottom sits exactly on the last data row is
*grown* to swallow the newly appended rows, which is what makes a weekly append
show up in the existing SUM.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

# Quoted string literals inside a formula, which must never be rewritten.
_STRING = re.compile(r'"(?:[^"]|"")*"')

_REF = re.compile(
    r"""
    (?P<sheet>(?:'(?:[^']|'')+'|[A-Za-z_\\][\w.]*)\s*!\s*)?   # optional sheet qualifier
    (?P<c1>\$?[A-Za-z]{1,3})(?P<r1>\$?\d{1,7})
    (?:\s*:\s*(?P<c2>\$?[A-Za-z]{1,3})(?P<r2>\$?\d{1,7}))?
    """,
    re.VERBOSE,
)

# Guard: do not treat a function name or a defined name as a reference.
_WORD_BEFORE = re.compile(r"[\w$!.]")


@dataclass(frozen=True)
class Ref:
    """A parsed A1 reference: one cell, or a rectangular range."""

    sheet: str | None
    col1: str
    row1: int
    row1_abs: bool
    col2: str | None = None
    row2: int | None = None
    row2_abs: bool = False

    @property
    def is_range(self) -> bool:
        return self.col2 is not None


def _sheet_name(qualifier: str | None) -> str | None:
    """Strip the trailing ``!`` and any surrounding quotes from a sheet qualifier."""
    if not qualifier:
        return None
    name = qualifier.rstrip().rstrip("!").rstrip()
    if name.startswith("'") and name.endswith("'"):
        name = name[1:-1].replace("''", "'")
    return name


def rewrite_refs(formula: str, transform) -> str:
    """Apply ``transform(Ref) -> (row1, row2) | None`` to every reference in a formula.

    String literals are skipped, and a match preceded by a word character is
    ignored so that ``LOG10`` or a defined name like ``Tax_2026`` is not mangled.
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula

    out: list[str] = []
    cursor = 0
    for literal in _STRING.finditer(formula):
        out.append(_rewrite_segment(formula[cursor : literal.start()], transform))
        out.append(literal.group(0))
        cursor = literal.end()
    out.append(_rewrite_segment(formula[cursor:], transform))
    return "".join(out)


def _rewrite_segment(segment: str, transform) -> str:
    result: list[str] = []
    cursor = 0
    for match in _REF.finditer(segment):
        start = match.start()
        if start > 0 and _WORD_BEFORE.match(segment[start - 1]):
            continue
        # A "reference" immediately followed by "(" is a function call, not a ref.
        tail = segment[match.end() : match.end() + 1]
        if tail == "(":
            continue

        r1_raw, r2_raw = match.group("r1"), match.group("r2")
        ref = Ref(
            sheet=_sheet_name(match.group("sheet")),
            col1=match.group("c1"),
            row1=int(r1_raw.lstrip("$")),
            row1_abs=r1_raw.startswith("$"),
            col2=match.group("c2"),
            row2=int(r2_raw.lstrip("$")) if r2_raw else None,
            row2_abs=bool(r2_raw and r2_raw.startswith("$")),
        )
        new = transform(ref)
        if new is None:
            continue
        new_r1, new_r2 = new
        if new_r1 == ref.row1 and new_r2 == ref.row2:
            continue

        rebuilt = f"{match.group('sheet') or ''}{ref.col1}{'$' if ref.row1_abs else ''}{new_r1}"
        if ref.is_range:
            rebuilt += f":{ref.col2}{'$' if ref.row2_abs else ''}{new_r2}"
        result.append(segment[cursor:start])
        result.append(rebuilt)
        cursor = match.end()
    result.append(segment[cursor:])
    return "".join(result)


def shift_for_insert(formula: str, *, sheet: str, target_sheet: str, at_row: int, count: int) -> str:
    """Update a formula for ``count`` rows inserted directly above ``at_row``.

    ``at_row`` is the first new row, so the last pre-existing data row is
    ``at_row - 1``. Four cases, in the order Excel applies them:

    * a reference below the insertion point moves down;
    * a range that straddles the insertion point stretches;
    * a range ending on the last data row is stretched too, so existing totals
      pick up the appended rows -- Excel would leave this one behind;
    * anything above the insertion point is left alone.

    ``sheet`` is where the formula lives, which is what unqualified references
    resolve against; ``target_sheet`` is where the rows are being inserted.
    """
    boundary = at_row - 1

    def transform(ref: Ref):
        ref_sheet = ref.sheet or sheet
        if ref_sheet != target_sheet:
            return None

        if not ref.is_range:
            return (ref.row1 + count, None) if ref.row1 >= at_row else None

        row1, row2 = ref.row1, ref.row2
        if row1 >= at_row:
            return row1 + count, row2 + count
        if row2 >= boundary:
            return row1, row2 + count
        return None

    return rewrite_refs(formula, transform)


def translate_rows(formula: str, delta: int) -> str:
    """Move every relative row reference in a formula by ``delta``.

    Used when copying a per-row formula (``=D11-E11``) down onto the rows being
    appended. Rows pinned with ``$`` stay put, matching a normal fill-down.
    """

    def transform(ref: Ref):
        row1 = ref.row1 if ref.row1_abs else ref.row1 + delta
        row2 = None
        if ref.is_range:
            row2 = ref.row2 if ref.row2_abs else ref.row2 + delta
        return row1, row2

    return rewrite_refs(formula, transform)


def shift_range_ref(ref: str, *, at_row: int, count: int) -> str:
    """Update a plain range string such as an autofilter or table ref."""
    return rewrite_refs("=" + ref, lambda r: _range_only(r, at_row, count))[1:]


def _range_only(ref: Ref, at_row: int, count: int):
    boundary = at_row - 1
    if not ref.is_range:
        return (ref.row1 + count, None) if ref.row1 >= at_row else None
    if ref.row1 >= at_row:
        return ref.row1 + count, ref.row2 + count
    if ref.row2 >= boundary:
        return ref.row1, ref.row2 + count
    return None
