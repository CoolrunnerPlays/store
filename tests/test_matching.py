"""The duplicate matcher, which is what failed in the audit run."""

from datetime import date
from decimal import Decimal

import pytest

from ledgerflow.matching import ExistingRow, Matcher, cents, coerce_date, normalize, similarity
from ledgerflow.models import Record


def existing(row=4, when=date(2026, 3, 7), description="AWS CLOUD SERVICES", amount="-832.10", reference=""):
    return ExistingRow(row=row, date=when, description=description, reference=reference,
                       amount_cents=cents(Decimal(amount)))


def incoming(when=date(2026, 3, 7), description="AWS CLOUD SERVICES", amount="-832.10", reference=""):
    return Record(kind="transaction", source="s.pdf", date=when, description=description,
                  amount=Decimal(amount), reference=reference)


def test_exact_repeat_is_a_duplicate():
    assert Matcher([existing()]).find(incoming()).status == "duplicate"


def test_a_reworded_description_is_still_a_duplicate():
    """The audit failure: a person types the payee their own way."""
    match = Matcher([existing(description="Aws - Cloud  Services")]).find(incoming())
    assert match.status == "duplicate"


def test_sign_convention_does_not_defeat_matching():
    """A book records an outgoing as positive; the parser reads it as negative."""
    match = Matcher([existing(amount="832.10")]).find(incoming(amount="-832.10"))
    assert match.status == "duplicate"


def test_a_date_stored_as_text_still_matches():
    """The other audit failure: dates typed as strings read as no date at all."""
    assert coerce_date("03/07/2026") == date(2026, 3, 7)
    assert coerce_date("2026-03-07") == date(2026, 3, 7)
    assert coerce_date(None) is None


def test_a_posting_date_a_day_out_is_held_back_not_added():
    match = Matcher([existing(when=date(2026, 3, 5))]).find(incoming(when=date(2026, 3, 7)))
    assert match.status == "probable_duplicate"


def test_a_matching_reference_settles_it():
    match = Matcher([existing(when=date(2020, 1, 1), description="something else", reference="INV-900")]).find(
        incoming(description="totally different", reference="INV-900")
    )
    assert match.status == "duplicate"


def test_a_genuinely_new_transaction_is_new():
    match = Matcher([existing()]).find(incoming(when=date(2026, 5, 4), description="NEW SUPPLIER", amount="-777.77"))
    assert match.status == "new"


def test_the_same_amount_on_a_far_off_date_is_new():
    match = Matcher([existing()]).find(incoming(when=date(2026, 9, 9)))
    assert match.status == "new"


def test_a_record_with_no_amount_is_never_silently_added():
    record = Record(kind="transaction", source="s.pdf", date=date(2026, 3, 7), description="SOMETHING")
    assert Matcher([existing()]).find(record).status == "unmatchable"


def test_a_repeat_inside_one_batch_is_caught():
    matcher = Matcher([])
    first = incoming()
    assert matcher.find(first).status == "new"
    matcher.add(first)
    assert matcher.find(incoming()).status == "duplicate"


def test_a_real_recurring_charge_on_a_different_date_is_new():
    """Same rent every month must not collapse into one row."""
    matcher = Matcher([existing(when=date(2026, 3, 1), description="OFFICE RENT", amount="-2900.00")])
    match = matcher.find(incoming(when=date(2026, 4, 1), description="OFFICE RENT", amount="-2900.00"))
    assert match.status == "new"


@pytest.mark.parametrize(
    "left,right",
    [
        ("CLIENT PAYMENT WESTFIELD LLC", "Client - Payment Westfield Llc"),
        ("PAYROLL RUN 03-2026", "Payroll  Run  03-2026"),
        ("OFFICE RENT MARCH", "Rent - office, March"),
    ],
)
def test_descriptions_that_mean_the_same_thing_score_high(left, right):
    assert similarity(left, right) >= 0.55


def test_unrelated_descriptions_score_low():
    assert similarity("UTILITIES CONSOLIDATED", "TESCO SUPERSTORE") < 0.55


def test_normalisation_strips_punctuation_and_case():
    assert normalize("  Client - Payment,  Westfield LLC ") == "CLIENT PAYMENT WESTFIELD LLC"
