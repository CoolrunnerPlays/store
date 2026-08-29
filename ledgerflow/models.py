"""Core record types shared by the extractors, the mapper and the appender."""

from __future__ import annotations

import hashlib
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from typing import Any, Literal

Confidence = Literal["ok", "uncertain"]


@dataclass
class Flag:
    """A value the extractor could not read with confidence.

    Flags never stop a run. They are collected and written to the Review Notes
    sheet so the numbers a human still has to check are visible in one place,
    instead of being silently guessed at.
    """

    source: str
    """File the value came from."""

    location: str
    """Where in that file, e.g. "page 2, line 14"."""

    field: str
    """Which field is doubtful, e.g. "amount"."""

    raw: str
    """The literal text as it was read."""

    reason: str
    """Why it could not be trusted."""

    row_key: str = ""
    """Fingerprint of the record this flag belongs to, blank if the whole record was dropped."""


@dataclass
class Record:
    """One row destined for the workbook.

    A bank transaction and an invoice both land here. The fields either side
    does not use stay ``None`` and the mapper simply leaves those columns alone.
    """

    kind: Literal["transaction", "invoice"]
    source: str
    date: date | None = None
    description: str = ""
    payee: str = ""
    amount: Decimal | None = None
    """Signed net amount: positive for money in, negative for money out."""
    deposit: Decimal | None = None
    withdrawal: Decimal | None = None
    balance: Decimal | None = None
    category: str = ""
    reference: str = ""
    """Invoice number, cheque number or bank reference."""
    vendor: str = ""
    tax: Decimal | None = None
    subtotal: Decimal | None = None
    total: Decimal | None = None
    line_items: list[str] = field(default_factory=list)
    page: int | None = None
    confidence: Confidence = "ok"
    extra: dict[str, Any] = field(default_factory=dict)

    def fingerprint(self) -> str:
        """Stable identity for this record, used to avoid re-adding it next week.

        Deliberately built from the *content* of the row rather than the file it
        arrived in: the same statement re-uploaded under a different name, or a
        month overlapping two PDFs, must not produce duplicate rows.
        """
        amount = self.amount if self.amount is not None else self.total
        parts = [
            self.kind,
            self.date.isoformat() if self.date else "",
            _norm(self.description or self.vendor),
            _norm(self.reference),
            f"{amount:.2f}" if amount is not None else "",
        ]
        return hashlib.sha1("|".join(parts).encode("utf-8")).hexdigest()[:16]


def _norm(text: str) -> str:
    """Collapse whitespace and case so trivial re-spacing does not defeat dedupe."""
    return " ".join(text.split()).upper()
