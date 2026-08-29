# LedgerFlow

Adds each week's bank statements and invoices to the spreadsheet you already keep.

It does not rebuild your file or hand you a new template. It opens the workbook you
maintain, works out its layout on its own, inserts the new rows inside your existing
table, and re-points your formulas so the totals you already have grow to cover them.

## What it protects

| | |
|---|---|
| Existing rows | Never touched. Verified cell by cell in the test suite. |
| Formatting | New rows inherit the styling, number formats and row height of the rows above. |
| Your formulas | `=SUM(D4:D11)` becomes `=SUM(D4:D23)`. Cross-sheet formulas on a Summary tab are updated too. |
| Per-row formulas | A calculated column such as `Net = Deposit - Withdrawal` is filled down onto the new rows. |
| Excel tables | A ListObject range, the autofilter, conditional formats, validations and merges all grow with the table. |
| Numbers | Written as numeric cells, so your own formulas keep working. Never as text. |
| Your original file | Never modified. A new `_updated.xlsx` copy is written. |

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

Every row carries a fingerprint built from its date, description, reference and amount.
Rows already in the workbook are skipped and reported, so overlapping statements or a
re-uploaded file cannot produce duplicates. This works even on a book that was filled in
by hand long before this tool existed, because the fingerprints of the existing rows are
recomputed from the sheet itself.

A hidden `_LedgerFlow` sheet records what has been added. Pass `--no-state` to leave it out.

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
tests/            69 tests, including a full round trip through the web API
```

## Tests

```bash
pip install -e ".[dev]"
pytest
```

The suite builds workbooks with the awkward parts a real book has — a title above the
header, per-row formulas, a totals block, an Excel table, a summary sheet pointing at the
data — and asserts that an append leaves all of it correct.
