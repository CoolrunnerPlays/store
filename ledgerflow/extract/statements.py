"""Turn a bank statement PDF into one Record per transaction.

Statements are column layouts, not sentences. The parser therefore locates the
statement's own header line, remembers where each column sits horizontally, and
assigns every amount on a row to a column by position. That is what separates a
deposit from a withdrawal reliably; falling back to "the last number on the
line" gets it wrong the moment a running balance appears.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from decimal import Decimal

from ..models import Flag, Record
from ..values import MONEY_TOKEN, infer_dayfirst, parse_date, parse_money
from .pdftext import Document, Line, Word

# Header words mapped to the role their column plays.
COLUMN_ROLES: dict[str, str] = {
    "date": "date",
    "posted": "date",
    "posting": "date",
    "transaction": "date",
    "value": "value_date",
    "description": "description",
    "details": "description",
    "particulars": "description",
    "narrative": "description",
    "payee": "description",
    "merchant": "description",
    "memo": "description",
    "reference": "reference",
    "ref": "reference",
    "cheque": "reference",
    "check": "reference",
    "category": "category",
    "type": "category",
    "debit": "withdrawal",
    "debits": "withdrawal",
    "withdrawal": "withdrawal",
    "withdrawals": "withdrawal",
    "payments": "withdrawal",
    "paid out": "withdrawal",
    "money out": "withdrawal",
    "credit": "deposit",
    "credits": "deposit",
    "deposit": "deposit",
    "deposits": "deposit",
    "paid in": "deposit",
    "money in": "deposit",
    "amount": "amount",
    "balance": "balance",
}

MONTH = r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?"

DATE_AT_START = re.compile(
    r"""^\s*(
        \d{1,2}[/-]\d{1,2}(?:[/-]\d{2,4})?     | # 03/14 or 03/14/2026
        \d{4}-\d{2}-\d{2}                       | # 2026-03-14
        \d{1,2}\s+MONTH(?:\s+\d{2,4})?          | # 14 Mar 2026
        MONTH\s+\d{1,2}(?:,?\s+\d{2,4})?         # Mar 14, 2026
    )""".replace("MONTH", MONTH),
    re.VERBOSE | re.IGNORECASE,
)

NOISE = re.compile(
    r"^\s*(page\s+\d+|continued|statement period|opening balance|closing balance|"
    r"balance (brought|carried) forward|total\b|subtotal\b)",
    re.IGNORECASE,
)


@dataclass
class ColumnAnchor:
    role: str
    x0: float
    x1: float

    def overlaps(self, word: Word) -> float:
        """How much of the word sits inside this column, as a 0-1 fraction."""
        span = max(word.x1 - word.x0, 0.1)
        return max(0.0, min(self.x1, word.x1) - max(self.x0, word.x0)) / span


def find_column_anchors(lines: list[Line]) -> list[ColumnAnchor]:
    """Locate the statement's header line and return the column positions.

    Returns an empty list when no header is recognisable, which sends the caller
    down the positional fallback path.
    """
    best: list[ColumnAnchor] = []
    for line in lines[:40]:
        anchors: list[ColumnAnchor] = []
        for word in line.words:
            key = re.sub(r"[^a-z ]", "", word.text.lower()).strip()
            role = COLUMN_ROLES.get(key)
            if role is None:
                continue
            if anchors and anchors[-1].role == role:
                anchors[-1] = ColumnAnchor(role, anchors[-1].x0, word.x1)
                continue
            anchors.append(ColumnAnchor(role, word.x0, word.x1))
        roles = {a.role for a in anchors}
        # A real header names a date plus at least one money column.
        if "date" in roles and roles & {"withdrawal", "deposit", "amount", "balance"} and len(anchors) > len(best):
            best = anchors
    return _widen(best)


def _widen(anchors: list[ColumnAnchor]) -> list[ColumnAnchor]:
    """Stretch each anchor to the midpoint of its neighbours.

    Header captions rarely line up with the digits beneath them -- an amount is
    right-aligned under a centred caption -- so the narrow caption box is
    expanded into a full column band before anything is matched against it.
    """
    anchors = sorted(anchors, key=lambda a: a.x0)
    widened: list[ColumnAnchor] = []
    for index, anchor in enumerate(anchors):
        left = anchors[index - 1].x1 if index else anchor.x0 - 40
        right = anchors[index + 1].x0 if index + 1 < len(anchors) else anchor.x1 + 60
        widened.append(ColumnAnchor(anchor.role, (left + anchor.x0) / 2, (anchor.x1 + right) / 2))
    return widened


def _role_for(word: Word, anchors: list[ColumnAnchor]) -> str | None:
    best_role, best_overlap = None, 0.0
    for anchor in anchors:
        overlap = anchor.overlaps(word)
        if overlap > best_overlap:
            best_role, best_overlap = anchor.role, overlap
    return best_role if best_overlap > 0.25 else None


MONEY_ROLES = {"withdrawal", "deposit", "amount", "balance"}

# A bare integer is only an amount when its column says so. Anywhere else it is
# far more likely to be a cheque number, an account number or a quantity.
CURRENCY_SHAPED = re.compile(r"[%s]|\d[.,]\d{2}\b|\(|\)|(?:CR|DR)$" % "$£€₹¥", re.IGNORECASE)


def _money_words(line: Line, anchors: list[ColumnAnchor]) -> list[Word]:
    """Words on a row that are genuinely currency amounts.

    Column position is the strongest signal: a number under the Deposit heading
    is a deposit even when written as a bare ``500``. A number sitting in the
    description column has to look like currency before it counts, which keeps
    cheque and reference numbers out of the arithmetic.
    """
    by_role: dict[str | None, list[Word]] = {}
    for word in line.words:
        if not MONEY_TOKEN.fullmatch(word.text) or parse_money(word.text) is None:
            continue
        role = _role_for(word, anchors) if anchors else None
        if role in MONEY_ROLES:
            by_role.setdefault(role, []).append(word)
        elif role is None and CURRENCY_SHAPED.search(word.text):
            by_role.setdefault(None, []).append(word)

    found: list[Word] = []
    for role, words in by_role.items():
        # A long description can spill a bare number into the column band next
        # to it. When a column holds several candidates, the ones written like
        # currency win and the rest fall back to being description text.
        if len(words) > 1:
            shaped = [w for w in words if CURRENCY_SHAPED.search(w.text)]
            if shaped:
                words = shaped
        found.extend(words)
    return sorted(found, key=lambda w: w.x0)


def parse_statement(doc: Document, *, source: str) -> tuple[list[Record], list[Flag]]:
    """Extract every transaction row from a statement document."""
    records: list[Record] = []
    flags: list[Flag] = []

    if not doc.has_text_layer:
        flags.append(
            Flag(
                source=source,
                location="whole file",
                field="document",
                raw="",
                reason="No text layer found - this looks like a scan. It needs visual review before its rows can be trusted.",
            )
        )
        return records, flags

    all_lines = [line for page in doc.pages for line in page.lines]
    date_samples = [m.group(1) for m in (DATE_AT_START.match(l.text) for l in all_lines) if m]
    dayfirst = infer_dayfirst(date_samples)
    year_hint = _infer_year(doc.text)

    anchors = find_column_anchors(all_lines)
    previous_balance: Decimal | None = None
    current: Record | None = None

    for line in all_lines:
        text = line.text.strip()
        if not text or NOISE.match(text):
            continue

        match = DATE_AT_START.match(text)
        if not match:
            if current is not None and not _money_words(line, anchors) and len(text) < 80:
                current.description = f"{current.description} {text}".strip()
            continue

        record, balance, line_flags = _parse_row(
            line, match, anchors, source=source, dayfirst=dayfirst, year_hint=year_hint,
            previous_balance=previous_balance,
        )
        flags.extend(line_flags)
        if record is None:
            continue
        if balance is not None:
            previous_balance = balance
        records.append(record)
        current = record

    return records, flags


def _infer_year(text: str) -> int | None:
    years = re.findall(r"\b(19|20)(\d{2})\b", text)
    if not years:
        return None
    return int(f"{years[0][0]}{years[0][1]}")


def _parse_row(
    line: Line,
    date_match: re.Match,
    anchors: list[ColumnAnchor],
    *,
    source: str,
    dayfirst: bool,
    year_hint: int | None,
    previous_balance: Decimal | None,
) -> tuple[Record | None, Decimal | None, list[Flag]]:
    """Parse a single dated line into a Record."""
    flags: list[Flag] = []
    where = f"page {line.page}, '{line.text[:60]}'"

    raw_date = date_match.group(1)
    when = parse_date(raw_date, dayfirst=dayfirst, year_hint=year_hint)
    if when is None:
        flags.append(Flag(source, where, "date", raw_date, "Date text could not be read unambiguously."))
        return None, None, flags

    money = _money_words(line, anchors)
    if not money:
        return None, None, flags

    buckets: dict[str, list[Word]] = {}
    for word in money:
        role = _role_for(word, anchors) if anchors else None
        buckets.setdefault(role or "unassigned", []).append(word)

    description = _description_text(line, date_match, money, anchors)

    deposit = _single(buckets.get("deposit"))
    withdrawal = _single(buckets.get("withdrawal"))
    balance = _single(buckets.get("balance"))
    amount = _single(buckets.get("amount"))
    leftover = buckets.get("unassigned", [])

    if not anchors or (deposit is None and withdrawal is None and amount is None):
        amount, balance, fallback_flags = _positional_fallback(
            line, money if not anchors else leftover, balance, source, where
        )
        flags.extend(fallback_flags)

    if withdrawal is not None:
        withdrawal = abs(withdrawal)
    if deposit is not None:
        deposit = abs(deposit)

    net = _net_amount(amount, deposit, withdrawal, balance, previous_balance)
    if net is None and amount is None and deposit is None and withdrawal is None:
        return None, None, flags

    if net is not None and deposit is None and withdrawal is None:
        deposit = net if net > 0 else None
        withdrawal = -net if net < 0 else None

    record = Record(
        kind="transaction",
        source=source,
        date=when,
        description=description,
        payee=description,
        amount=net,
        deposit=deposit,
        withdrawal=withdrawal,
        balance=balance,
        reference=_reference(description),
        page=line.page,
    )

    if net is not None and amount is not None and balance is not None and previous_balance is not None:
        implied = balance - previous_balance
        if abs(abs(implied) - abs(net)) > Decimal("0.01"):
            record.confidence = "uncertain"
            flags.append(
                Flag(
                    source, where, "amount", line.text[:80],
                    f"Amount {net} does not reconcile with the balance change of {implied}.",
                    row_key=record.fingerprint(),
                )
            )

    return record, balance, flags


def _single(words: list[Word] | None) -> Decimal | None:
    """One amount per column per row; anything else is not a clean read."""
    if not words or len(words) > 1:
        return None
    return parse_money(words[0].text)


def _net_amount(
    amount: Decimal | None,
    deposit: Decimal | None,
    withdrawal: Decimal | None,
    balance: Decimal | None,
    previous_balance: Decimal | None,
) -> Decimal | None:
    """Signed amount: positive in, negative out.

    When the statement uses a single Amount column with no sign, the direction is
    recovered from the movement in the running balance rather than assumed.
    """
    if deposit is not None and withdrawal is not None:
        return deposit - withdrawal
    if deposit is not None:
        return deposit
    if withdrawal is not None:
        return -withdrawal
    if amount is None:
        return None
    if amount != 0:
        return amount
    if balance is not None and previous_balance is not None:
        delta = balance - previous_balance
        if abs(abs(delta) - abs(amount)) <= Decimal("0.01"):
            return amount.copy_abs() if delta > 0 else -amount.copy_abs()
    return amount


def _positional_fallback(
    line: Line, money: list[Word], balance: Decimal | None, source: str, where: str
) -> tuple[Decimal | None, Decimal | None, list[Flag]]:
    """Last resort when the statement has no recognisable header.

    Two trailing amounts mean amount-then-balance; one means just an amount.
    Three or more is genuinely ambiguous and is flagged rather than guessed.
    """
    flags: list[Flag] = []
    values = [parse_money(w.text) for w in money]
    values = [v for v in values if v is not None]
    if not values:
        return None, balance, flags
    if len(values) == 1:
        return values[0], balance, flags
    if len(values) == 2:
        return values[0], values[1], flags
    flags.append(
        Flag(source, where, "amount", line.text[:80],
             f"{len(values)} amounts on one row with no column header to tell them apart.")
    )
    return values[0], values[-1], flags


def _description_text(line: Line, date_match: re.Match, money: list[Word], anchors: list[ColumnAnchor]) -> str:
    """Everything on the row that is neither the leading date nor an amount."""
    money_ids = {id(w) for w in money}
    date_text = date_match.group(1).split()
    parts: list[str] = []
    skipped = 0
    for word in line.words:
        if id(word) in money_ids:
            continue
        if skipped < len(date_text) and word.text in date_text:
            skipped += 1
            continue
        parts.append(word.text)
    return " ".join(parts).strip(" -–")


_REF_RE = re.compile(r"\b(?:REF|REFERENCE|CHQ|CHEQUE|CHECK|TXN|TRN)\b[:#\s]*([A-Z0-9-]{3,})\b", re.IGNORECASE)


def _reference(description: str) -> str:
    match = _REF_RE.search(description)
    return match.group(1) if match else ""
