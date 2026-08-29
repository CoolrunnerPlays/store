"""Turn an invoice PDF into a single Record.

Invoices are label-driven rather than columnar: the useful values sit next to
words like "Invoice No" or "Total Due". The parser looks for those labels, and
where an invoice states a subtotal, a tax and a total, it checks that the three
agree before trusting them.
"""

from __future__ import annotations

import re
from decimal import Decimal

from ..models import Flag, Record
from ..values import MONEY_TOKEN, infer_dayfirst, parse_date, parse_money
from .pdftext import Document, Line

INVOICE_NO = re.compile(
    r"\b(?:invoice|inv|bill|document)\s*(?:number|no\.?|num|#)?\s*[:#]\s*([A-Za-z0-9][A-Za-z0-9/_-]{2,})"
    r"|\b(?:invoice|inv)\s*(?:number|no\.?|#)\s+([A-Za-z0-9][A-Za-z0-9/_-]{2,})",
    re.IGNORECASE,
)

DATE_LABEL = re.compile(
    r"\b(invoice\s*date|date\s*of\s*issue|issue\s*date|bill\s*date|dated|date)\b\s*[:#]?\s*(.{4,30})",
    re.IGNORECASE,
)

EXCLUDE_DATE_LABEL = re.compile(r"\b(due|payment|period|delivery|service)\s*date\b|\bdue\b", re.IGNORECASE)

TOTAL_LABELS = [
    ("total", re.compile(r"\b(?:total\s*due|amount\s*due|balance\s*due|grand\s*total|invoice\s*total|total)\b", re.I)),
    ("tax", re.compile(r"\b(?:vat|gst|hst|sales\s*tax|tax)\b", re.I)),
    ("subtotal", re.compile(r"\b(?:sub[\s-]*total|net\s*(?:amount|total)?|amount\s*before\s*tax)\b", re.I)),
]

ITEM_HEADER = re.compile(r"\b(description|item|details|service|product|qty|quantity)\b", re.IGNORECASE)

VENDOR_NOISE = re.compile(r"^\s*(invoice|tax invoice|bill|statement|receipt|proforma)\s*$", re.IGNORECASE)


def parse_invoice(doc: Document, *, source: str) -> tuple[list[Record], list[Flag]]:
    """Extract one invoice Record from a document."""
    flags: list[Flag] = []
    if not doc.has_text_layer:
        flags.append(
            Flag(
                source=source,
                location="whole file",
                field="document",
                raw="",
                reason="No text layer found - this looks like a scan or a photo. It needs visual review.",
            )
        )
        return [], flags

    lines = [line for page in doc.pages for line in page.lines]
    text = "\n".join(line.text for line in lines)

    number = _invoice_number(text)
    vendor = _vendor(lines)
    when, date_flag = _invoice_date(lines, source)
    if date_flag:
        flags.append(date_flag)

    amounts = _labelled_amounts(lines)
    subtotal, tax, total = amounts.get("subtotal"), amounts.get("tax"), amounts.get("total")
    items = _line_items(lines)

    if total is None:
        # Some invoices only print a single figure; take the largest amount as
        # the total but say so, because that is an inference rather than a read.
        candidates = [parse_money(w.text) for line in lines for w in line.words if MONEY_TOKEN.fullmatch(w.text)]
        candidates = [c for c in candidates if c is not None]
        if candidates:
            total = max(candidates)
            flags.append(
                Flag(source, "totals block", "total", str(total),
                     "No 'Total' label found; used the largest amount on the invoice.")
            )

    record = Record(
        kind="invoice",
        source=source,
        date=when,
        vendor=vendor,
        description="; ".join(items) if items else vendor,
        reference=number,
        subtotal=subtotal,
        tax=tax,
        total=total,
        amount=-total if total is not None else None,
        line_items=items,
        page=1,
    )

    if subtotal is not None and tax is not None and total is not None:
        if abs((subtotal + tax) - total) > Decimal("0.02"):
            record.confidence = "uncertain"
            flags.append(
                Flag(source, "totals block", "total", f"{subtotal} + {tax} != {total}",
                     "Subtotal plus tax does not equal the stated total.",
                     row_key=record.fingerprint())
            )

    for field_name, value in (("invoice number", number), ("invoice date", when), ("total", total)):
        if not value:
            record.confidence = "uncertain"
            flags.append(
                Flag(source, "whole file", field_name, "", f"No {field_name} could be located.",
                     row_key=record.fingerprint())
            )

    return [record], flags


def _invoice_number(text: str) -> str:
    match = INVOICE_NO.search(text)
    if not match:
        return ""
    return (match.group(1) or match.group(2) or "").strip(" .,:")


def _vendor(lines: list[Line]) -> str:
    """The issuing company, taken as the most prominent line at the top of page 1.

    Font size beats position here: an invoice usually sets the vendor name in
    the largest type on the page, above or beside a smaller address block.
    """
    head = [l for l in lines[:12] if l.text.strip() and not VENDOR_NOISE.match(l.text.strip())]
    if not head:
        return ""
    largest = max(head, key=lambda l: l.size)
    if largest.size > 0:
        return largest.text.strip()
    return head[0].text.strip()


def _invoice_date(lines: list[Line], source: str) -> tuple[object, Flag | None]:
    """The invoice's own date, never its due date."""
    samples = [l.text for l in lines]
    dayfirst = infer_dayfirst(samples)
    fallback: str | None = None

    for line in lines:
        text = line.text
        if EXCLUDE_DATE_LABEL.search(text):
            continue
        match = DATE_LABEL.search(text)
        if not match:
            continue
        raw = match.group(2).strip(" :#")
        when = parse_date(raw, dayfirst=dayfirst)
        if when is not None:
            return when, None
        fallback = raw

    if fallback:
        return None, Flag(source, "date line", "invoice date", fallback, "Date text next to the label could not be read.")
    return None, None


def _labelled_amounts(lines: list[Line]) -> dict[str, Decimal]:
    """Pick out subtotal, tax and total from the summary block.

    Labels are tested most-specific first, and each label is only taken once, so
    a 'Total' further down cannot overwrite the 'Total Due' already found.
    """
    found: dict[str, Decimal] = {}
    for line in lines:
        text = line.text
        amounts = [parse_money(w.text) for w in line.words if MONEY_TOKEN.fullmatch(w.text)]
        amounts = [a for a in amounts if a is not None]
        if not amounts:
            continue
        for key, pattern in TOTAL_LABELS:
            if key in found or not pattern.search(text):
                continue
            found[key] = amounts[-1]
            break
    return found


def _line_items(lines: list[Line]) -> list[str]:
    """Descriptions of the billed items, between the item header and the totals."""
    start = None
    for index, line in enumerate(lines):
        if ITEM_HEADER.search(line.text) and len(line.text) < 90:
            start = index + 1
            break
    if start is None:
        return []

    items: list[str] = []
    for line in lines[start:]:
        text = line.text.strip()
        if not text:
            continue
        if any(pattern.search(text) for _, pattern in TOTAL_LABELS):
            break
        description = _strip_trailing_amounts(line)
        if description and not description.isdigit():
            items.append(description)
    return items[:20]


def _strip_trailing_amounts(line: Line) -> str:
    """Drop the qty/unit/amount columns from the right of an item row.

    Only the unbroken run of numeric tokens at the end of the line is removed,
    so a quantity written inside the description -- "Asset licensing (12
    months)" -- survives intact.
    """
    words = list(line.words)
    while words and MONEY_TOKEN.fullmatch(words[-1].text) and parse_money(words[-1].text) is not None:
        words.pop()
    return " ".join(w.text for w in words).strip(" -\u2013:")
