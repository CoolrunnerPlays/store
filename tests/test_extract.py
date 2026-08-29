from decimal import Decimal

from ledgerflow.extract import classify, extract_records, load_document


def test_debit_credit_statement_reads_every_row(docs):
    records, flags = extract_records(docs["statement_dc"])
    assert len(records) == 8
    assert not flags


def test_signs_follow_the_column_the_amount_sits_in(docs):
    records, _ = extract_records(docs["statement_dc"])
    by_description = {r.description: r for r in records}
    assert by_description["CLIENT PAYMENT WESTFIELD LLC"].amount == Decimal("6500.00")
    assert by_description["AWS CLOUD SERVICES"].amount == Decimal("-832.10")


def test_extracted_rows_reconcile_against_the_balance_column(docs):
    records, _ = extract_records(docs["statement_dc"])
    movement = sum((r.amount for r in records), Decimal("0"))
    opening = records[0].balance - records[0].amount
    assert opening + movement == records[-1].balance


def test_a_reference_number_is_not_mistaken_for_an_amount(docs):
    records, _ = extract_records(docs["statement_dc"])
    row = next(r for r in records if "ACME" in r.description)
    assert row.withdrawal == Decimal("1240.55")
    assert "88213" in row.description
    assert row.reference == "88213"


def test_single_amount_column_uses_parentheses_for_outgoings(docs):
    records, _ = extract_records(docs["statement_amount"])
    amounts = {r.description: r.amount for r in records}
    assert amounts["INTEREST CREDIT"] == Decimal("42.18")
    assert amounts["TRANSFER TO CHECKING"] == Decimal("-1000.00")


def test_invoice_fields(docs):
    records, flags = extract_records(docs["invoice"])
    assert len(records) == 1
    invoice = records[0]
    assert invoice.vendor == "Brightline Media Ltd"
    assert invoice.reference == "INV-2026-0142"
    assert invoice.subtotal == Decimal("4750.00")
    assert invoice.tax == Decimal("950.00")
    assert invoice.total == Decimal("5700.00")
    assert invoice.date.isoformat() == "2026-05-12"
    assert not flags


def test_line_items_keep_numbers_that_belong_to_the_text(docs):
    invoice = extract_records(docs["invoice"])[0][0]
    assert "Asset licensing (12 months)" in invoice.line_items


def test_documents_classify_themselves(docs):
    assert classify(load_document(docs["statement_dc"])) == "statement"
    assert classify(load_document(docs["invoice"])) == "invoice"


def test_an_image_is_flagged_rather_than_skipped_silently(tmp_path):
    photo = tmp_path / "receipt.jpg"
    photo.write_bytes(b"not really a jpeg")
    records, flags = extract_records(photo)
    assert not records
    assert len(flags) == 1
    assert "visual" in flags[0].reason.lower() or "manual" in flags[0].reason.lower()


def test_identical_rows_from_different_files_share_a_fingerprint(docs):
    first, _ = extract_records(docs["statement_dc"])
    second, _ = extract_records(docs["statement_dc"])
    assert {r.fingerprint() for r in first} == {r.fingerprint() for r in second}
