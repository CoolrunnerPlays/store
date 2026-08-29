"""Route each uploaded file to the right extractor.

Statements and invoices are told apart by their own vocabulary rather than by
filename, because filenames are unreliable and users drop a whole week's folder
in at once.
"""

from __future__ import annotations

import json
import re
from pathlib import Path

from ..models import Flag, Record
from .invoices import parse_invoice
from .pdftext import Document, read_pdf
from .statements import parse_statement

IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".webp", ".gif", ".tif", ".tiff", ".heic"}

STATEMENT_HINTS = re.compile(
    r"\b(statement|account\s*(?:number|no|summary)|sort\s*code|opening\s*balance|closing\s*balance|"
    r"balance\s*(?:brought|carried)\s*forward|withdrawals?|deposits?|transactions?|iban|routing)\b",
    re.IGNORECASE,
)
INVOICE_HINTS = re.compile(
    r"\b(invoice|bill\s*to|remit\s*to|due\s*date|purchase\s*order|subtotal|vat|gst|tax\s*invoice|"
    r"payment\s*terms|amount\s*due)\b",
    re.IGNORECASE,
)


def classify(doc: Document) -> str:
    """Return "statement" or "invoice" for a document."""
    text = doc.text
    statement_score = len(STATEMENT_HINTS.findall(text))
    invoice_score = len(INVOICE_HINTS.findall(text))
    # Many dated rows is the strongest statement signal there is.
    dated_rows = len(re.findall(r"^\s*\d{1,4}[/-]\d{1,2}[/-]\d{1,4}", text, re.MULTILINE))
    if dated_rows >= 4:
        statement_score += dated_rows // 2
    return "statement" if statement_score > invoice_score else "invoice"


def load_document(path: str | Path) -> Document | None:
    """Read a PDF, or return None for a file this tool cannot read as text."""
    path = Path(path)
    if path.suffix.lower() == ".pdf":
        return read_pdf(str(path))
    return None


def extract_records(path: str | Path, *, kind: str | None = None) -> tuple[list[Record], list[Flag]]:
    """Extract records from one file.

    ``kind`` forces "statement" or "invoice" when the caller already knows;
    otherwise the document classifies itself.
    """
    path = Path(path)
    source = path.name

    if path.suffix.lower() in IMAGE_SUFFIXES:
        return [], [
            Flag(
                source=source,
                location="whole file",
                field="document",
                raw="",
                reason="Image file: no text to read automatically. Add it with a manual-entry file, "
                       "or have the numbers read off the picture and confirmed before they go in.",
            )
        ]

    if path.suffix.lower() == ".json":
        return _load_manual(path)

    doc = load_document(path)
    if doc is None:
        return [], [
            Flag(source, "whole file", "document", "", f"Unsupported file type '{path.suffix}'.")
        ]

    resolved = kind or classify(doc)
    if resolved == "statement":
        return parse_statement(doc, source=source)
    return parse_invoice(doc, source=source)


def _load_manual(path: Path) -> tuple[list[Record], list[Flag]]:
    """Read hand-entered records, the escape hatch for scans and photos.

    The file is a JSON list of objects using the same field names as Record, so
    anything that cannot be read automatically can still be appended through the
    same dedupe, formula and formatting machinery as everything else.
    """
    from datetime import date as date_cls
    from decimal import Decimal

    payload = json.loads(path.read_text())
    if isinstance(payload, dict):
        payload = payload.get("records", [])

    records: list[Record] = []
    flags: list[Flag] = []
    money_fields = {"amount", "deposit", "withdrawal", "balance", "tax", "subtotal", "total"}

    for index, item in enumerate(payload, start=1):
        try:
            fields = dict(item)
            when = fields.pop("date", None)
            for key in list(fields):
                if key in money_fields and fields[key] is not None:
                    fields[key] = Decimal(str(fields[key]))
            allowed = {f for f in Record.__dataclass_fields__ if f not in {"date", "kind", "source"}}
            fields = {k: v for k, v in fields.items() if k in allowed}
            records.append(
                Record(
                    kind=item.get("kind", "transaction"),
                    source=item.get("source", path.name),
                    date=date_cls.fromisoformat(when) if when else None,
                    **fields,
                )
            )
        except Exception as error:  # a bad hand-written row must not stop the rest
            flags.append(Flag(path.name, f"entry {index}", "record", json.dumps(item)[:120], str(error)))
    return records, flags
