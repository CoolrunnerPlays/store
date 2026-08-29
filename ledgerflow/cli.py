"""Command line entry point.

Three verbs, matching how the work actually happens: look at the book once,
append to it every week, or open the browser version for someone who would
rather drag files onto a page.
"""

from __future__ import annotations

import argparse
import sys
from decimal import Decimal
from pathlib import Path

from .append import append_records, choose_sheet
from .extract import extract_records
from .introspect import TableInfo
from .mapping import Profile, suggest_profile
from .models import Flag, Record

DOC_SUFFIXES = {".pdf", ".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff", ".json"}


def _money(value: Decimal) -> str:
    return f"{value:,.2f}"


def _collect(paths: list[str]) -> list[Path]:
    """Expand any directories in the arguments into the documents inside them."""
    files: list[Path] = []
    for raw in paths:
        path = Path(raw)
        if path.is_dir():
            files.extend(sorted(p for p in path.iterdir() if p.suffix.lower() in DOC_SUFFIXES))
        elif path.exists():
            files.append(path)
        else:
            print(f"  ! {raw} does not exist, skipping", file=sys.stderr)
    return files


def _print_table(name: str, table: TableInfo, profile: Profile, marker: str = "") -> None:
    print(f"\nSheet '{name}'{marker}")
    print(f"  headers on row {table.header_row}, data rows {table.first_data_row}-{table.last_data_row} "
          f"({table.row_count} rows)")
    if table.total_rows:
        print(f"  totals row(s): {', '.join(str(r) for r in table.total_rows)}")
    if table.excel_table:
        print(f"  Excel table: {table.excel_table}")
    print("  columns:")
    for column in table.columns:
        mapped = profile.field_for(column.letter)
        if column.is_computed:
            note = f"formula, filled down as {column.formula_template}"
        elif mapped:
            note = f"<- {mapped}"
        else:
            note = "left alone"
        print(f"    {column.letter:>3}  {column.header[:26]:26s} {column.inferred_type:6s}  {note}")


def cmd_inspect(args) -> int:
    """Report what the tool sees in a workbook and save a mapping profile."""
    sheet, tables = choose_sheet(args.workbook)
    target = args.sheet or sheet

    for name, table in tables.items():
        profile = suggest_profile(table)
        marker = "   <- this one will be used" if name == target else ""
        _print_table(name, table, profile, marker)

    profile = suggest_profile(tables[target])
    unmapped = [f for f in ("date", "description", "amount") if f not in profile.columns.values()]
    if "amount" in unmapped and {"deposit", "withdrawal", "total"} & set(profile.columns.values()):
        unmapped.remove("amount")
    if unmapped:
        print(f"\n  note: no column found for {', '.join(unmapped)}. "
              f"Edit the profile if one of the columns above should take it.")

    destination = Path(args.profile) if args.profile else Path(args.workbook).with_suffix(".profile.json")
    profile.save(destination)
    print(f"\nProfile written to {destination}")
    print("Edit it if a column is mapped to the wrong field, then run 'ledgerflow add'.")
    return 0


def cmd_add(args) -> int:
    """Extract from the given documents and append to the workbook."""
    files = _collect(args.documents)
    if not files:
        print("No documents to read.", file=sys.stderr)
        return 1

    records: list[Record] = []
    flags: list[Flag] = []
    print("Reading documents:")
    for path in files:
        found, issues = extract_records(path, kind=args.kind)
        records.extend(found)
        flags.extend(issues)
        kinds = {r.kind for r in found}
        label = "/".join(sorted(kinds)) if kinds else "nothing"
        print(f"  {path.name:44s} {len(found):>3} records ({label})"
              f"{f', {len(issues)} to review' if issues else ''}")

    if args.profile:
        profile = Profile.load(args.profile)
    else:
        sheet, tables = choose_sheet(args.workbook)
        target = args.sheet or sheet
        if target not in tables:
            print(f"Sheet {target!r} has no detectable table.", file=sys.stderr)
            return 1
        profile = suggest_profile(tables[target])
        print(f"\nUsing sheet '{profile.sheet}' with an auto-detected mapping "
              f"({', '.join(f'{k}->{v}' for k, v in profile.columns.items())}).")
        print("Run 'ledgerflow inspect' first if you want to fix the mapping before writing.")

    if args.sheet:
        profile.sheet = args.sheet

    result = append_records(
        args.workbook,
        records,
        profile,
        output_path=args.out,
        flags=flags,
        dry_run=args.dry_run,
        track_state=not args.no_state,
    )

    print("\n" + "=" * 62)
    print(f"  Invoices processed          {len(result.invoices)}")
    print(f"  Bank transactions extracted {len(result.transactions)}")
    print(f"  Money in                    {_money(result.money_in)}")
    print(f"  Money out                   {_money(result.money_out)}")
    print(f"  Invoice value               {_money(result.invoice_total)}")
    print(f"  Net change added            {_money(result.total_added)}")
    if result.skipped_duplicates:
        print(f"  Already in the book         {len(result.skipped_duplicates)} (skipped)")
    if result.skipped_other_kind:
        print(f"  Not for this sheet          {len(result.skipped_other_kind)} (skipped)")
    print(f"  Flagged for review          {len(result.flags)}")
    print("=" * 62)

    if result.flags:
        print("\nNeeds a human eye:")
        for flag in result.flags[:20]:
            print(f"  - {flag.source} ({flag.location}) {flag.field}: {flag.reason}")
        if len(result.flags) > 20:
            print(f"  ... and {len(result.flags) - 20} more, all listed on the Review Notes sheet.")

    if args.dry_run:
        print("\nDry run: nothing was written.")
        if result.added:
            print(f"Would insert {len(result.added)} rows at row {result.first_new_row} of '{result.sheet}'.")
    elif result.added:
        print(f"\nWrote {len(result.added)} rows into '{result.sheet}' starting at row {result.first_new_row}.")
        print(f"{result.formulas_updated} formulas re-pointed so the existing totals cover them.")
        print(f"Saved: {result.output}")
    else:
        print("\nNothing new to add - everything in these documents is already in the book.")
    return 0


def cmd_web(args) -> int:
    from .web.app import run

    run(host=args.host, port=args.port, debug=args.debug)
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="ledgerflow",
        description="Append bank statements and invoices to a spreadsheet you already keep.",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    inspect = sub.add_parser("inspect", help="show what the tool sees in a workbook and save a mapping profile")
    inspect.add_argument("workbook")
    inspect.add_argument("--sheet", help="sheet to map, if not the one auto-selected")
    inspect.add_argument("--profile", help="where to write the profile JSON")
    inspect.set_defaults(func=cmd_inspect)

    add = sub.add_parser("add", help="extract documents and append them to the workbook")
    add.add_argument("workbook")
    add.add_argument("documents", nargs="+", help="PDF files, images, manual-entry JSON, or folders of them")
    add.add_argument("--sheet", help="sheet to append to")
    add.add_argument("--profile", help="mapping profile to use instead of auto-detecting")
    add.add_argument("--out", help="output file (default: <workbook>_updated.xlsx)")
    add.add_argument("--kind", choices=["statement", "invoice"], help="force how documents are read")
    add.add_argument("--dry-run", action="store_true", help="report what would happen and write nothing")
    add.add_argument("--no-state", action="store_true", help="do not add the hidden tracking sheet")
    add.set_defaults(func=cmd_add)

    web = sub.add_parser("web", help="run the drag-and-drop browser version")
    web.add_argument("--host", default="127.0.0.1")
    web.add_argument("--port", type=int, default=5000)
    web.add_argument("--debug", action="store_true")
    web.set_defaults(func=cmd_web)

    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        return args.func(args)
    except (ValueError, FileNotFoundError) as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
