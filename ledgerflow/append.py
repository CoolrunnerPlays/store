"""Append new records into an existing workbook without disturbing it.

The rules this module keeps, in priority order:

1. Nothing already in the file changes meaning. Existing rows, styles, column
   widths, freeze panes and formulas come out the way they went in.
2. Totals stay correct. Rows go *inside* the table, and every formula in the
   workbook is re-pointed so the existing SUMs cover the new rows.
3. Numbers are numbers. Amounts are written as numeric cells so the user's own
   formulas keep working.
4. Running it twice does not double up. Records carry a content fingerprint and
   already-present rows are skipped.
"""

from __future__ import annotations

from copy import copy
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any, Iterable

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .formulas import shift_for_insert, shift_range_ref, translate_rows
from .introspect import TableInfo, describe_sheet, describe_workbook
from .mapping import MONEY_FIELDS, Profile, value_for
from .models import Flag, Record

STATE_SHEET = "_LedgerFlow"
REVIEW_SHEET = "Review Notes"


@dataclass
class AppendResult:
    """What one run did, in the terms the summary needs."""

    output: Path
    sheet: str
    added: list[Record] = field(default_factory=list)
    skipped_duplicates: list[Record] = field(default_factory=list)
    skipped_other_kind: list[Record] = field(default_factory=list)
    flags: list[Flag] = field(default_factory=list)
    first_new_row: int = 0
    formulas_updated: int = 0

    @property
    def invoices(self) -> list[Record]:
        return [r for r in self.added if r.kind == "invoice"]

    @property
    def transactions(self) -> list[Record]:
        return [r for r in self.added if r.kind == "transaction"]

    @property
    def money_in(self) -> Decimal:
        return sum((r.deposit for r in self.added if r.deposit), Decimal("0"))

    @property
    def money_out(self) -> Decimal:
        return sum((r.withdrawal for r in self.added if r.withdrawal), Decimal("0"))

    @property
    def invoice_total(self) -> Decimal:
        return sum((r.total for r in self.invoices if r.total), Decimal("0"))

    @property
    def total_added(self) -> Decimal:
        """Sum of what was appended, using each record's own headline figure."""
        total = Decimal("0")
        for record in self.added:
            value = record.amount if record.amount is not None else record.total
            if value is not None:
                total += value
        return total


def append_records(
    workbook_path: str | Path,
    records: Iterable[Record],
    profile: Profile,
    *,
    output_path: str | Path | None = None,
    flags: list[Flag] | None = None,
    track_state: bool = True,
    dry_run: bool = False,
) -> AppendResult:
    """Append records to the sheet named by ``profile`` and save a new file."""
    workbook_path = Path(workbook_path)
    output = Path(output_path) if output_path else workbook_path.with_name(
        f"{workbook_path.stem}_updated{workbook_path.suffix}"
    )

    wb = load_workbook(workbook_path, data_only=False, keep_vba=workbook_path.suffix.lower() == ".xlsm")
    if profile.sheet not in wb.sheetnames:
        raise ValueError(f"Sheet {profile.sheet!r} is not in {workbook_path.name}")
    ws = wb[profile.sheet]

    table = describe_sheet(ws)
    if table is None:
        raise ValueError(f"No data table could be found on sheet {profile.sheet!r}")

    result = AppendResult(output=output, sheet=profile.sheet, flags=list(flags or []))

    wanted, result.skipped_other_kind = _split_by_kind(records, profile)
    known = _existing_fingerprints(wb, ws, table, profile) if profile.skip_duplicates else set()
    fresh: list[Record] = []
    for record in wanted:
        key = record.fingerprint()
        if key in known:
            result.skipped_duplicates.append(record)
            continue
        known.add(key)
        fresh.append(record)

    if profile.sort_by_date:
        fresh.sort(key=lambda r: (r.date or date.max, r.description))

    result.added = fresh
    if dry_run or (not fresh and not result.flags):
        result.first_new_row = _insertion_row(table)
        return result

    if not fresh:
        # Nothing new to append, but there are things to look at. The notes are
        # worth a file on their own; the ledger itself is untouched.
        result.first_new_row = _insertion_row(table)
        _write_review_notes(wb, result)
        wb.save(output)
        return result

    at_row = _insertion_row(table)
    count = len(fresh)
    result.first_new_row = at_row

    style_source = table.last_data_row if table.last_data_row > table.header_row else None

    ws.insert_rows(at_row, count)
    result.formulas_updated = _repoint_workbook(wb, profile.sheet, at_row, count)
    _repair_ranges(ws, at_row, count)

    for offset, record in enumerate(fresh):
        row = at_row + offset
        if style_source is not None:
            _copy_row_style(ws, style_source, row, table)
        _write_row(ws, row, record, table, profile, style_source)

    _write_review_notes(wb, result)
    if track_state:
        _write_state(wb, [r.fingerprint() for r in fresh])

    wb.save(output)
    return result


def _insertion_row(table: TableInfo) -> int:
    """First row of the new block: directly under the last row of real data."""
    return max(table.last_data_row, table.header_row) + 1


def _split_by_kind(records: Iterable[Record], profile: Profile) -> tuple[list[Record], list[Record]]:
    wanted, rejected = [], []
    for record in records:
        (wanted if record.kind in profile.kinds else rejected).append(record)
    return wanted, rejected


def _existing_fingerprints(wb, ws: Worksheet, table: TableInfo, profile: Profile) -> set[str]:
    """Fingerprints of rows already in the book.

    Read from the tracking sheet where one exists, and always recomputed from the
    visible rows as well, so a book that was filled in by hand before this tool
    existed still dedupes correctly on the first run.
    """
    known: set[str] = set()

    if STATE_SHEET in wb.sheetnames:
        for row in wb[STATE_SHEET].iter_rows(min_row=2, max_col=1, values_only=True):
            if row[0]:
                known.add(str(row[0]))

    by_field = {field: letter for letter, field in profile.columns.items()}
    for row in range(table.first_data_row, table.last_data_row + 1):
        record = _row_to_record(ws, row, by_field, profile)
        if record is not None:
            known.add(record.fingerprint())
    return known


def _row_to_record(ws: Worksheet, row: int, by_field: dict[str, str], profile: Profile) -> Record | None:
    """Rebuild enough of a Record from an existing row to fingerprint it."""

    def cell(field_name: str):
        letter = by_field.get(field_name)
        if not letter:
            return None
        value = ws[f"{letter}{row}"].value
        return None if isinstance(value, str) and value.startswith("=") else value

    when = cell("date")
    if isinstance(when, datetime):
        when = when.date()
    description = cell("description") or cell("payee") or cell("vendor") or ""
    reference = cell("reference") or ""

    amount = _as_decimal(cell("amount"))
    if amount is None:
        deposit, withdrawal = _as_decimal(cell("deposit")), _as_decimal(cell("withdrawal"))
        if deposit is not None or withdrawal is not None:
            amount = (deposit or Decimal(0)) - (withdrawal or Decimal(0))
    total = _as_decimal(cell("total"))

    if when is None and not description and amount is None and total is None:
        return None

    kind = "invoice" if profile.kinds == ["invoice"] else "transaction"
    return Record(
        kind=kind,
        source="existing",
        date=when if isinstance(when, date) else None,
        description=str(description),
        reference=str(reference),
        amount=amount,
        total=total,
    )


def _as_decimal(value: Any) -> Decimal | None:
    if value is None or isinstance(value, str):
        return None
    try:
        return Decimal(str(value)).quantize(Decimal("0.01"))
    except Exception:
        return None


def _repoint_workbook(wb, target_sheet: str, at_row: int, count: int) -> int:
    """Rewrite every formula in the book for the inserted rows.

    openpyxl moves cells but never touches formula text, so without this pass a
    totals row that moved down would still be summing the rows it used to sit
    under. Runs over all sheets because a summary tab commonly points here.
    """
    updated = 0
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                value = cell.value
                if not isinstance(value, str) or not value.startswith("="):
                    continue
                rewritten = shift_for_insert(
                    value, sheet=sheet.title, target_sheet=target_sheet, at_row=at_row, count=count
                )
                if rewritten != value:
                    cell.value = rewritten
                    updated += 1

    for name, defined in list(getattr(wb, "defined_names", {}).items()):
        try:
            attr_text = defined.attr_text
        except AttributeError:
            continue
        if not attr_text:
            continue
        rewritten = shift_for_insert(
            "=" + attr_text, sheet=target_sheet, target_sheet=target_sheet, at_row=at_row, count=count
        )[1:]
        if rewritten != attr_text:
            defined.attr_text = rewritten
            updated += 1
    return updated


def _repair_ranges(ws: Worksheet, at_row: int, count: int) -> None:
    """Grow the sheet features that openpyxl leaves pointing at the old rows.

    Excel tables, the autofilter, conditional formats, validations and merges all
    store their own A1 ranges. Left alone, banding stops halfway down the table
    and dropdowns go missing on the new rows.
    """
    for table in list(getattr(ws, "tables", {}).values()):
        table.ref = shift_range_ref(table.ref, at_row=at_row, count=count)

    if ws.auto_filter and ws.auto_filter.ref:
        ws.auto_filter.ref = shift_range_ref(ws.auto_filter.ref, at_row=at_row, count=count)

    if getattr(ws, "print_area", None):
        areas = ws.print_area if isinstance(ws.print_area, (list, tuple)) else [ws.print_area]
        ws.print_area = [shift_range_ref(str(a).replace("$", ""), at_row=at_row, count=count) for a in areas]

    for rule_range in list(getattr(ws.conditional_formatting, "_cf_rules", {}).keys()):
        rules = ws.conditional_formatting._cf_rules[rule_range]
        new_ref = " ".join(
            shift_range_ref(part, at_row=at_row, count=count) for part in str(rule_range.sqref).split()
        )
        if new_ref != str(rule_range.sqref):
            del ws.conditional_formatting._cf_rules[rule_range]
            rule_range.sqref = new_ref
            ws.conditional_formatting._cf_rules[rule_range] = rules

    for validation in list(getattr(ws, "data_validations", None).dataValidation if ws.data_validations else []):
        parts = [shift_range_ref(str(p), at_row=at_row, count=count) for p in str(validation.sqref).split()]
        validation.sqref = " ".join(parts)

    merged = [str(m) for m in ws.merged_cells.ranges]
    for ref in merged:
        shifted = shift_range_ref(ref, at_row=at_row, count=count)
        if shifted != ref:
            ws.unmerge_cells(ref)
            ws.merge_cells(shifted)


def _copy_row_style(ws: Worksheet, source_row: int, target_row: int, table: TableInfo) -> None:
    """Give a new row the same look as the row it follows."""
    for col in range(table.first_col, table.last_col + 1):
        source = ws.cell(row=source_row, column=col)
        target = ws.cell(row=target_row, column=col)
        target._style = copy(source._style)
    height = ws.row_dimensions[source_row].height
    if height is not None:
        ws.row_dimensions[target_row].height = height


def _write_row(
    ws: Worksheet,
    row: int,
    record: Record,
    table: TableInfo,
    profile: Profile,
    style_source: int | None,
) -> None:
    """Fill one new row: mapped values, and the sheet's own per-row formulas."""
    for column in table.columns:
        cell = ws.cell(row=row, column=column.index)

        if column.is_computed:
            # Carry the sheet's own calculation down, the way filling down would.
            cell.value = translate_rows(column.formula_template, row - table.last_data_row)
            if style_source is None:
                cell.number_format = column.number_format
            continue

        field_name = profile.field_for(column.letter)
        if field_name is None:
            continue

        value = value_for(record, field_name)
        if value is None:
            continue

        if field_name in MONEY_FIELDS and isinstance(value, Decimal):
            cell.value = float(value)
            if style_source is None:
                cell.number_format = column.number_format
        elif isinstance(value, date):
            cell.value = value
            cell.number_format = profile.date_format or column.number_format or "yyyy-mm-dd"
        else:
            cell.value = str(value)


def _write_review_notes(wb, result: AppendResult) -> None:
    """Write the flags to their own sheet, newest run at the top of its block."""
    if not result.flags:
        return

    if REVIEW_SHEET in wb.sheetnames:
        ws = wb[REVIEW_SHEET]
        start = ws.max_row + 2
    else:
        ws = wb.create_sheet(REVIEW_SHEET)
        start = 1
        headers = ["Run", "Source file", "Where", "Field", "Value as read", "Why it needs checking"]
        for index, name in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=index, value=name)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="B45309")
        widths = {"A": 20, "B": 28, "C": 40, "D": 16, "E": 24, "F": 60}
        for letter, width in widths.items():
            ws.column_dimensions[letter].width = width
        ws.freeze_panes = "A2"
        start = 2

    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    for offset, flag in enumerate(result.flags):
        row = start + offset
        for index, value in enumerate(
            [stamp, flag.source, flag.location, flag.field, flag.raw, flag.reason], start=1
        ):
            cell = ws.cell(row=row, column=index, value=value)
            cell.alignment = Alignment(vertical="top", wrap_text=index == 6)


def _write_state(wb, fingerprints: list[str]) -> None:
    """Record what has been added, so next week's run knows what is already in."""
    if STATE_SHEET in wb.sheetnames:
        ws = wb[STATE_SHEET]
    else:
        ws = wb.create_sheet(STATE_SHEET)
        ws["A1"] = "fingerprint"
        ws["B1"] = "added"
        ws.sheet_state = "hidden"

    stamp = datetime.now().isoformat(timespec="seconds")
    row = ws.max_row + 1
    for offset, fingerprint in enumerate(fingerprints):
        ws.cell(row=row + offset, column=1, value=fingerprint)
        ws.cell(row=row + offset, column=2, value=stamp)


def choose_sheet(workbook_path: str | Path) -> tuple[str, dict[str, TableInfo]]:
    """Pick the most ledger-like sheet in a workbook, and report all candidates."""
    wb = load_workbook(workbook_path, data_only=False, read_only=False)
    tables = {name: info for name, info in describe_workbook(wb).items() if info.looks_like_table}
    if not tables:
        raise ValueError("No table with headers and data rows was found in this workbook.")
    best = max(tables.items(), key=lambda item: item[1].score)[0]
    return best, tables
