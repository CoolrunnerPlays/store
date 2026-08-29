"""Parsing of the money and date strings that appear in statements and invoices.

Everything here returns ``None`` rather than a guess when the text is not
unambiguous. Callers turn that ``None`` into a Flag.
"""

from __future__ import annotations

import re
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

from dateutil import parser as dateparser

_CURRENCY = "$£€₹¥"
_MONEY_RE = re.compile(
    r"""
    (?P<open>\()?                 # accounting negative
    \s*[%s]?\s*                   # optional currency symbol
    (?P<sign>[-+])?\s*
    (?P<num>\d{1,3}(?:[,\s.]\d{3})*(?:[.,]\d{1,2})?|\d+(?:[.,]\d{1,2})?)
    \s*(?P<close>\))?
    \s*(?P<suffix>CR|DR|-)?       # trailing credit/debit marker
    """
    % _CURRENCY,
    re.VERBOSE | re.IGNORECASE,
)

# A bare token that is money-shaped, used when scanning a statement line.
MONEY_TOKEN = re.compile(r"\(?[%s]?-?\d[\d,\s.]*\d?(?:[.,]\d{2})?\)?(?:\s?(?:CR|DR))?" % _CURRENCY, re.IGNORECASE)


def parse_money(text: str | None, *, eu_format: bool = False) -> Decimal | None:
    """Return the Decimal value of a money string, or None if it is not readable.

    Handles ``$1,234.56``, ``(1,234.56)``, ``1234.56-``, ``1.234,56`` (with
    ``eu_format``) and trailing ``CR``/``DR`` markers. Returns None for empty
    text, for OCR debris, and for anything with more than two decimal places,
    which in practice means the token was not a currency amount at all.
    """
    if text is None:
        return None
    raw = str(text).strip()
    if not raw or raw in {"-", "--", "."}:
        return None

    match = _MONEY_RE.fullmatch(raw)
    if not match:
        return None

    num = match.group("num")
    if eu_format:
        num = num.replace(".", "").replace(" ", "").replace(",", ".")
    else:
        # Thousands separators may be commas or spaces; the decimal point is a dot.
        if "," in num and "." in num:
            num = num.replace(",", "")
        elif "," in num:
            # A lone comma is a decimal comma only when it is followed by 1-2 digits.
            head, _, tail = num.rpartition(",")
            num = f"{head.replace(',', '')}.{tail}" if len(tail) in (1, 2) else num.replace(",", "")
        num = num.replace(" ", "")

    try:
        value = Decimal(num)
    except InvalidOperation:
        return None

    negative = bool(match.group("open") and match.group("close"))
    negative |= match.group("sign") == "-"
    suffix = (match.group("suffix") or "").upper()
    negative |= suffix in {"DR", "-"}
    if suffix == "CR":
        negative = False
    return -value if negative else value


def looks_like_money(text: str) -> bool:
    """True when a token is shaped like a currency amount."""
    return parse_money(text) is not None


def parse_date(text: str | None, *, dayfirst: bool = False, year_hint: int | None = None) -> date | None:
    """Return a date, or None when the text is not an unambiguous date.

    ``dayfirst`` follows the statement's own convention, which the statement
    parser infers once per document rather than guessing per line. ``year_hint``
    supplies the year for statements that print ``12 Mar`` with no year.
    """
    if not text:
        return None
    raw = " ".join(str(text).split())
    if not raw:
        return None
    if isinstance(text, (datetime, date)):
        return text.date() if isinstance(text, datetime) else text

    # Reject bare numbers: "1234" is not a date even though dateutil will take it.
    if re.fullmatch(r"\d{1,4}", raw):
        return None

    default = datetime(year_hint or datetime.now().year, 1, 1)
    try:
        parsed = dateparser.parse(raw, dayfirst=dayfirst, fuzzy=False, default=default)
    except (ValueError, OverflowError, TypeError):
        return None
    if parsed is None:
        return None
    return parsed.date()


def infer_dayfirst(samples: list[str]) -> bool:
    """Decide whether a document writes dates day-first.

    Looks for a sample where the first component exceeds 12, which can only be a
    day. Falls back to month-first (US convention) when nothing is decisive.
    """
    for sample in samples:
        match = re.match(r"\s*(\d{1,2})[/-](\d{1,2})[/-]", sample)
        if not match:
            continue
        first, second = int(match.group(1)), int(match.group(2))
        if first > 12 and second <= 12:
            return True
        if second > 12 and first <= 12:
            return False
    return False
