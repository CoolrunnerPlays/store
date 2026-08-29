"""Deciding whether a record is already in the workbook.

The first version of this compared an exact hash of date + description +
reference + amount. That is far too strict for a book a person maintains: the
same payment is typed with different wording, a date is stored as the text
"03/07/2026" instead of a real date, a reference is left blank. Any one of those
made a row look new, and the audit run appended a complete second copy of a
ledger that was already full.

Matching now leans on the two things that come from the document rather than
from anyone's typing -- the date and the amount -- and uses the description only
to confirm. Nothing is written unless a record is genuinely unmatched.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from difflib import SequenceMatcher
from typing import Iterable, Literal

from .models import Record

MatchStatus = Literal["new", "duplicate", "probable_duplicate", "unmatchable"]

DEFAULT_DATE_TOLERANCE_DAYS = 3
"""A posting date and a value date commonly differ by a day or two."""

STRONG_SIMILARITY = 0.55
"""Enough to confirm a same-date, same-amount hit is the same transaction."""

WEAK_SIMILARITY = 0.5
"""Enough to confirm a near-date, same-amount hit."""

# Bank noise that carries no identity: card fragments, terminal ids, boilerplate.
NOISE_TOKENS = {
    "POS", "PURCHASE", "PAYMENT", "TRANSACTION", "TXN", "TRN", "REF", "CARD",
    "DEBIT", "CREDIT", "VISA", "MASTERCARD", "AUTH", "PENDING", "ONLINE",
    "TRANSFER", "XX", "XXXX", "THE", "LTD", "LIMITED", "INC", "LLC", "CO",
}


def normalize(text: str) -> str:
    """Strip a description down to the part that identifies the payee."""
    if not text:
        return ""
    cleaned = re.sub(r"[^A-Za-z0-9 ]+", " ", str(text)).upper()
    return " ".join(cleaned.split())


def tokens(text: str) -> set[str]:
    """Identifying words of a description, with boilerplate and digits removed."""
    words = set(normalize(text).split())
    return {w for w in words if w not in NOISE_TOKENS and not w.isdigit() and len(w) > 1}


def similarity(left: str, right: str) -> float:
    """How alike two descriptions are, on a 0-1 scale.

    Combines a character-level ratio with a word-overlap ratio and takes the
    better of the two, so both a re-spelled description and a re-ordered one
    still match.
    """
    a, b = normalize(left), normalize(right)
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    sequence = SequenceMatcher(None, a, b).ratio()
    left_tokens, right_tokens = tokens(left), tokens(right)
    if left_tokens and right_tokens:
        overlap = len(left_tokens & right_tokens) / len(left_tokens | right_tokens)
        containment = len(left_tokens & right_tokens) / min(len(left_tokens), len(right_tokens))
        return max(sequence, overlap, containment * 0.9)
    return sequence


def cents(value: Decimal | float | int | None) -> int | None:
    """Absolute amount in whole cents, which is how amounts are compared.

    Absolute because a book may record an outgoing as a positive number in a
    Withdrawal column while the parser reads it as a negative net amount. The
    magnitude is the reliable part; the sign is a convention.
    """
    if value is None:
        return None
    try:
        return int((Decimal(str(value)).copy_abs() * 100).quantize(Decimal("1")))
    except Exception:
        return None


@dataclass
class ExistingRow:
    """A row already in the workbook, reduced to what identifies it."""

    row: int
    date: date | None
    description: str
    reference: str
    amount_cents: int | None

    @property
    def is_usable(self) -> bool:
        """Whether this row carries enough to be matched against at all."""
        return self.amount_cents is not None or (self.date is not None and bool(self.description))


@dataclass
class MatchResult:
    status: MatchStatus
    row: int | None = None
    reason: str = ""
    score: float = 0.0

    @property
    def is_duplicate(self) -> bool:
        return self.status in ("duplicate", "probable_duplicate")


@dataclass
class Matcher:
    """Indexes the rows already present and answers "have I seen this before?"."""

    rows: list[ExistingRow] = field(default_factory=list)
    date_tolerance_days: int = DEFAULT_DATE_TOLERANCE_DAYS
    _by_date_amount: dict[tuple[str, int], list[ExistingRow]] = field(default_factory=dict, repr=False)
    _by_amount: dict[int, list[ExistingRow]] = field(default_factory=dict, repr=False)
    _by_reference: dict[str, list[ExistingRow]] = field(default_factory=dict, repr=False)

    def __post_init__(self) -> None:
        for row in self.rows:
            if row.amount_cents is not None:
                self._by_amount.setdefault(row.amount_cents, []).append(row)
                if row.date:
                    self._by_date_amount.setdefault((row.date.isoformat(), row.amount_cents), []).append(row)
            reference = normalize(row.reference)
            if reference:
                self._by_reference.setdefault(reference, []).append(row)

    @property
    def usable_rows(self) -> int:
        return sum(1 for row in self.rows if row.is_usable)

    def add(self, record: Record, row: int = 0) -> None:
        """Remember a record just accepted, so a later duplicate in the same batch is caught."""
        existing = to_existing(record, row)
        self.rows.append(existing)
        if existing.amount_cents is not None:
            self._by_amount.setdefault(existing.amount_cents, []).append(existing)
            if existing.date:
                self._by_date_amount.setdefault(
                    (existing.date.isoformat(), existing.amount_cents), []
                ).append(existing)
        reference = normalize(existing.reference)
        if reference:
            self._by_reference.setdefault(reference, []).append(existing)

    def find(self, record: Record) -> MatchResult:
        """Classify a record against everything already known."""
        amount = cents(record.amount if record.amount is not None else record.total)
        reference = normalize(record.reference)

        # An invoice number is a unique document id: a match on it settles the question.
        if reference:
            for candidate in self._by_reference.get(reference, []):
                if amount is None or candidate.amount_cents is None or candidate.amount_cents == amount:
                    return MatchResult("duplicate", candidate.row, f"same reference {record.reference}", 1.0)

        if amount is None:
            return MatchResult(
                "unmatchable", None,
                "No amount could be read, so this cannot be checked against the rows already present.",
            )

        # Same day, same amount: almost always the same transaction.
        if record.date:
            for candidate in self._by_date_amount.get((record.date.isoformat(), amount), []):
                score = similarity(record.description, candidate.description)
                if score >= STRONG_SIMILARITY or not candidate.description or not record.description:
                    return MatchResult("duplicate", candidate.row, "same date and amount", score)
                return MatchResult(
                    "probable_duplicate", candidate.row,
                    f"same date and amount, but the wording differs (row {candidate.row}: "
                    f"{candidate.description[:40]!r})",
                    score,
                )

        # Same amount within a few days: a posting-date difference.
        best: MatchResult | None = None
        for candidate in self._by_amount.get(amount, []):
            if not record.date or not candidate.date:
                continue
            gap = abs((record.date - candidate.date).days)
            if gap > self.date_tolerance_days:
                continue
            score = similarity(record.description, candidate.description)
            if score >= WEAK_SIMILARITY and (best is None or score > best.score):
                best = MatchResult(
                    "probable_duplicate", candidate.row,
                    f"same amount, {gap} day(s) apart, wording matches row {candidate.row}",
                    score,
                )
        if best is not None:
            return best

        return MatchResult("new")


def to_existing(record: Record, row: int = 0) -> ExistingRow:
    amount = record.amount if record.amount is not None else record.total
    return ExistingRow(
        row=row,
        date=record.date,
        description=record.description or record.vendor,
        reference=record.reference,
        amount_cents=cents(amount),
    )


def coerce_date(value) -> date | None:
    """Read a date cell that may hold a real date or a typed string.

    A book kept by hand very often stores dates as text. Treating those as
    "no date" is what let a full ledger look empty to the matcher.
    """
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str) and value.strip():
        from .values import infer_dayfirst, parse_date

        return parse_date(value, dayfirst=infer_dayfirst([value]))
    return None


def build_matcher(rows: Iterable[ExistingRow], *, date_tolerance_days: int = DEFAULT_DATE_TOLERANCE_DAYS) -> Matcher:
    return Matcher(rows=list(rows), date_tolerance_days=date_tolerance_days)
