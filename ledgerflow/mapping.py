"""Match extracted fields to the columns of the user's own spreadsheet.

Nobody should have to describe their layout to the tool, so column meaning is
guessed from the header text. The guess is then written to a profile file, which
is the thing that makes week two a single command: the mapping is reviewed once
and reused every week after.
"""

from __future__ import annotations

import json
import re
from dataclasses import asdict, dataclass, field
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Any

from .introspect import TableInfo
from .models import Record

# Record field -> header words that mean it. Longer phrases are matched first so
# "invoice date" beats a bare "date", and "total amount" beats "amount".
SYNONYMS: dict[str, list[str]] = {
    "date": ["transactiondate", "invoicedate", "postingdate", "posteddate", "date", "day"],
    "description": ["description", "details", "particulars", "narrative", "memo", "item", "notes", "note"],
    "payee": ["payee", "merchant", "paidto", "supplier", "customer", "name"],
    "vendor": ["vendor", "company", "supplier", "billedby", "vendorname", "companyname"],
    "category": ["category", "type", "class", "account", "costcentre", "costcenter", "expensetype"],
    "reference": ["reference", "ref", "invoiceno", "invoicenumber", "invoice", "chequeno", "checkno",
                  "documentno", "docno", "transactionid", "txnid", "number", "no"],
    "deposit": ["deposit", "deposits", "credit", "credits", "moneyin", "paidin", "in", "income", "received"],
    "withdrawal": ["withdrawal", "withdrawals", "debit", "debits", "moneyout", "paidout", "out",
                   "expense", "expenses", "spend", "payment", "payments"],
    "amount": ["amount", "netamount", "value", "sum", "net"],
    "balance": ["balance", "runningbalance", "closingbalance"],
    "tax": ["tax", "vat", "gst", "hst", "salestax", "taxamount"],
    "subtotal": ["subtotal", "netofvat", "amountbeforetax", "goods"],
    "total": ["total", "totalamount", "grandtotal", "amountdue", "totaldue", "invoicetotal", "gross"],
    "source": ["source", "sourcefile", "file", "document", "attachment"],
}

# Fields that only ever carry money, so their columns must end up numeric.
MONEY_FIELDS = {"deposit", "withdrawal", "amount", "balance", "tax", "subtotal", "total"}


def _norm(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(text).lower())


def guess_field(header: str) -> str | None:
    """Best matching Record field for a column header, or None."""
    key = _norm(header)
    if not key:
        return None
    best: tuple[int, str] | None = None
    for field_name, words in SYNONYMS.items():
        for word in words:
            if key == word:
                return field_name
            if word in key or key in word:
                score = len(word)
                if best is None or score > best[0]:
                    best = (score, field_name)
    return best[1] if best else None


@dataclass
class Profile:
    """A saved decision about where each piece of data goes.

    Kept next to the workbook as JSON so the mapping is auditable and editable;
    the tool never silently changes where a value lands between runs.
    """

    sheet: str
    columns: dict[str, str] = field(default_factory=dict)
    """Column letter -> Record field name."""
    date_format: str = ""
    kinds: list[str] = field(default_factory=lambda: ["transaction", "invoice"])
    """Which record kinds belong on this sheet."""
    sort_by_date: bool = True
    skip_duplicates: bool = True

    def field_for(self, letter: str) -> str | None:
        return self.columns.get(letter)

    def to_json(self) -> str:
        return json.dumps(asdict(self), indent=2)

    @classmethod
    def load(cls, path: str | Path) -> "Profile":
        return cls(**json.loads(Path(path).read_text()))

    def save(self, path: str | Path) -> Path:
        path = Path(path)
        path.write_text(self.to_json())
        return path


def suggest_profile(table: TableInfo, *, kinds: list[str] | None = None) -> Profile:
    """Build a starting profile by reading the table's own headers.

    Computed columns are deliberately left unmapped: a column that carries a
    formula should keep carrying that formula on the new rows, not be overwritten
    with a value.
    """
    columns: dict[str, str] = {}
    taken: set[str] = set()
    date_format = ""

    for column in sorted(table.columns, key=lambda c: -len(_norm(c.header))):
        if column.is_computed or not column.header:
            continue
        guess = guess_field(column.header)
        if guess is None or guess in taken:
            continue
        columns[column.letter] = guess
        taken.add(guess)
        if guess == "date" and column.inferred_type == "date":
            date_format = column.number_format

    resolved_kinds = kinds or _kinds_for(taken)
    return Profile(
        sheet=table.sheet,
        columns=dict(sorted(columns.items())),
        date_format=date_format,
        kinds=resolved_kinds,
    )


def _kinds_for(fields: set[str]) -> list[str]:
    """Infer whether a sheet is for transactions, invoices, or both."""
    invoice_only = {"vendor", "tax", "subtotal", "total"} & fields
    bank_only = {"deposit", "withdrawal", "balance"} & fields
    if invoice_only and not bank_only:
        return ["invoice"]
    if bank_only and not invoice_only:
        return ["transaction"]
    return ["transaction", "invoice"]


def value_for(record: Record, field_name: str) -> Any:
    """The value a record contributes to one mapped column.

    Falls back between related fields so a book with a single Amount column still
    fills from an invoice total, and an invoice sheet's Vendor column still fills
    from a transaction's payee.
    """
    value = getattr(record, field_name, None)

    if field_name == "amount" and value is None:
        value = record.total if record.kind == "invoice" else None
    elif field_name == "total" and value is None:
        value = record.amount
    elif field_name == "vendor" and not value:
        value = record.payee or record.description
    elif field_name == "payee" and not value:
        value = record.vendor or record.description
    elif field_name == "description" and not value:
        value = record.vendor
    elif field_name == "withdrawal" and value is None and record.kind == "invoice":
        value = record.total
    elif field_name == "source":
        value = record.source

    if isinstance(value, list):
        value = "; ".join(str(v) for v in value)
    if isinstance(value, Decimal):
        return value
    if isinstance(value, date):
        return value
    if value in (None, ""):
        return None
    return value
