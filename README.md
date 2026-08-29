# LedgerFlow

Adds each week's bank statements and invoices to the spreadsheet you already keep.

It does not rebuild your file or hand you a new template. It opens the workbook you
maintain, works out its layout on its own, inserts the new rows inside your existing
table, and re-points your formulas so the totals you already have grow to cover them.

## What it protects

| | |
|---|---|
| Existing rows | Never touched. Verified cell by cell in the written file, on every run. |
| Formatting | New rows inherit the styling, number formats and row height of the rows above. |
| Your formulas | `=SUM(D4:D11)` becomes `=SUM(D4:D23)`. Cross-sheet formulas on a Summary tab are updated too. |
| Per-row formulas | A calculated column such as `Net = Deposit - Withdrawal` is filled down onto the new rows. |
| Excel tables | A ListObject range, the autofilter, conditional formats, validations and merges all grow with the table. |
| Numbers | Written as numeric cells, so your own formulas keep working. Never as text. |
| Your original file | Never modified. A new `_updated.xlsx` copy is written. |
| Row heights | Shifted with their rows on insert, which openpyxl does not do by itself. |
| Every write | Re-read and compared against the original before you get the file. |

## Install

```bash
pip install -e .
```

Python 3.10+. Dependencies: openpyxl, pdfplumber, python-dateutil, Flask.

## The weekly routine

**In the browser** — the easier way:

```bash
ledgerflow web
```

Open http://127.0.0.1:5000, drag in your spreadsheet and this week's PDFs, check the
rows it found, click add, download the file. Nothing leaves your machine.

**On the command line** — once, to see how your book was read:

```bash
ledgerflow inspect Accounts2026.xlsx
```

This prints the detected header row, the data range, the totals rows and which record
field will fill each column, then writes `Accounts2026.profile.json`. Edit that file if
anything is mapped to the wrong column.

Then every week:

```bash
ledgerflow add Accounts2026.xlsx ~/statements/week-32/ --profile Accounts2026.profile.json
```

Point it at individual files or a whole folder. Add `--dry-run` to see what would happen
without writing anything.

## Running it twice is safe

A record is compared against what is already in the sheet on the two things that
come from the document rather than from anyone's typing: **the date and the amount**.
The description is used only to confirm, and is compared loosely, so all of these
still count as the same transaction:

- the payee typed a person's own way (`AWS CLOUD SERVICES` vs `Aws - Cloud Services`);
- a date stored as the text `03/07/2026` instead of a real date cell;
- an outgoing recorded as a positive number in a Withdrawal column while the parser
  reads it as a negative net amount;
- a posting date a day or two off (held back and reported, not added).

An invoice number matching settles the question on its own.

Every record ends up in one of four buckets, and all four are shown before anything
is written:

| Verdict | What happens |
|---|---|
| **New** | Appended. |
| **Duplicate** | Already in the file. Excluded, with the row number it matched. |
| **Probable duplicate** | Same amount, date a few days out. Excluded and reported, so you can add it by hand if it really is separate. |
| **Unreadable** | No amount could be read, so it cannot be checked. Excluded and flagged. |

## The pre-commit safety check

**If there is nothing new, nothing is written.** A run that finds only duplicates
stops before the workbook is opened for modification, reports

```
0 new transactions found (100% duplicate history)
```

and leaves your file byte-identical. No output copy is produced either — an
identical file would only be something else to keep track of.

The run also aborts, untouched, when the sheet has rows but none of them can be read
as a date and an amount. Being unable to compare is a reason to stop, not a reason to
append.

## Every write is verified

After writing, the output is re-opened and compared against the original cell by cell:
value, font family, size, weight, colour, fill, all four borders, alignment, number
format, row heights and column widths. Rows above the insertion point are compared
against themselves; rows below against their new positions. If anything that should
have been left alone has changed, the write is rejected with `PreservationError` and
you get no file rather than a damaged one.

This exists because openpyxl's `insert_rows` moves cells but leaves `row_dimensions`
bound to the old row numbers, so every row below an insert silently inherits the wrong
height. LedgerFlow shifts them itself, and the verifier proves it on your actual file
rather than asking you to take it on trust.

## When something cannot be read

Nothing is guessed. Anything ambiguous is written to a **Review Notes** sheet in the
output file, with the source file, where in it, the text as read, and why it needs a look:

- a scanned PDF or a photo with no text layer;
- an amount that does not reconcile with the running balance;
- an invoice whose subtotal plus tax does not equal its stated total;
- a row with several amounts and no column headings to tell them apart.

For a scan or a photo, read the numbers off it and pass them in as JSON. They go through
the same dedupe, formatting and formula machinery as everything else:

```json
[
  {
    "kind": "invoice",
    "date": "2026-05-12",
    "vendor": "Brightline Media Ltd",
    "reference": "INV-2026-0142",
    "subtotal": 4750.00, "tax": 950.00, "total": 5700.00
  }
]
```

```bash
ledgerflow add Accounts2026.xlsx manual-entries.json
```

## How statements are read

Bank statements are column layouts, so the parser finds the statement's own header line,
records where each column sits horizontally, and assigns every amount to a column by
position. That is what separates a deposit from a withdrawal reliably — taking "the last
number on the line" gets it wrong as soon as a running balance appears. Where a statement
uses one signed Amount column, direction is recovered from the movement in the balance.

If a statement has no recognisable header, it falls back to reading the trailing amounts
positionally, and flags any row where that is genuinely ambiguous.

## Layout

```
ledgerflow/
  introspect.py   detects the header row, data range, totals rows and computed columns
  formulas.py     rewrites A1 references when rows are inserted
  mapping.py      matches record fields to your column headers; the profile file
  append.py       the append engine: dedupe, insert, style, re-point, review notes
  values.py       money and date parsing that returns None instead of guessing
  extract/        statement, invoice and document-routing parsers
  web/            the browser front end
  cli.py          inspect / add / web
tests/            109 tests, including the audit regressions and a web round trip
```

## Tests

```bash
pip install -e ".[dev]"
pytest
```

The suite builds workbooks with the awkward parts a real book has — a title above the
header, per-row formulas, a totals block, custom row heights, an Excel table, a summary
sheet pointing at the data — and asserts that an append leaves all of it correct.

`tests/test_audit_regression.py` pins the failures found in the audit run: a workbook
already holding every transaction must gain nothing and must not be rewritten, across
four ways a person's own book differs from the parser's reading. Two of its tests
deliberately reintroduce the old bugs to prove the verifier catches them.
