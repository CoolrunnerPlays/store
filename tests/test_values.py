from datetime import date
from decimal import Decimal

import pytest

from ledgerflow.values import infer_dayfirst, parse_date, parse_money


@pytest.mark.parametrize(
    "text,expected",
    [
        ("$1,234.56", Decimal("1234.56")),
        ("(1,234.56)", Decimal("-1234.56")),
        ("1234.56-", Decimal("-1234.56")),
        ("45.00 CR", Decimal("45.00")),
        ("45.00 DR", Decimal("-45.00")),
        ("-99.90", Decimal("-99.90")),
        ("1 234.56", Decimal("1234.56")),
        ("500", Decimal("500")),
    ],
)
def test_parses_money_shapes(text, expected):
    assert parse_money(text) == expected


@pytest.mark.parametrize("text", ["abc", "", "-", "12/03/2025", "3.14159"])
def test_refuses_to_guess_at_non_money(text):
    assert parse_money(text) is None


def test_european_decimal_comma():
    assert parse_money("1.234,56", eu_format=True) == Decimal("1234.56")


def test_dayfirst_changes_the_reading():
    assert parse_date("03/04/2025", dayfirst=True) == date(2025, 4, 3)
    assert parse_date("03/04/2025", dayfirst=False) == date(2025, 3, 4)


def test_bare_number_is_not_a_date():
    assert parse_date("1234") is None


def test_dayfirst_inferred_from_an_impossible_month():
    assert infer_dayfirst(["13/04/2025"]) is True
    assert infer_dayfirst(["04/13/2025"]) is False
    assert infer_dayfirst(["01/02/2025"]) is False
