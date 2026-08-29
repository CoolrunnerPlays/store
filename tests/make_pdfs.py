"""Generate sample statement and invoice PDFs for the test suite.

Real bank statements cannot be committed, so the suite builds documents with
the same shapes: a debit/credit/balance layout, a single signed-amount layout,
and a standard invoice.
"""

from __future__ import annotations

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

FONT = "Helvetica"


def _row(c, y, cells):
    for text, x, align in cells:
        if align == "r":
            c.drawRightString(x, y, text)
        else:
            c.drawString(x, y, text)


def statement_debit_credit(path):
    """Date | Description | Withdrawal | Deposit | Balance."""
    c = canvas.Canvas(str(path), pagesize=letter)
    c.setFont(FONT + "-Bold", 11)
    c.drawString(60, 740, "NORTHBANK BUSINESS CHECKING")
    c.setFont(FONT, 9)
    c.drawString(60, 726, "Statement period: 01 March 2026 to 31 March 2026")
    c.drawString(60, 714, "Account 0088 4412")

    c.setFont(FONT + "-Bold", 9)
    _row(c, 690, [("Date", 60, "l"), ("Description", 120, "l"),
                  ("Withdrawal", 400, "r"), ("Deposit", 470, "r"), ("Balance", 545, "r")])
    c.setFont(FONT, 9)

    rows = [
        ("03/02/2026", "OPENING PURCHASE ACME SUPPLIES REF 88213", "1,240.55", "", "18,759.45"),
        ("03/04/2026", "CLIENT PAYMENT WESTFIELD LLC", "", "6,500.00", "25,259.45"),
        ("03/07/2026", "AWS CLOUD SERVICES", "832.10", "", "24,427.35"),
        ("03/11/2026", "PAYROLL RUN 03-2026", "9,410.00", "", "15,017.35"),
        ("03/15/2026", "CLIENT PAYMENT NORTHGATE CO", "", "3,250.75", "18,268.10"),
        ("03/19/2026", "OFFICE RENT MARCH", "2,900.00", "", "15,368.10"),
        ("03/22/2026", "REFUND VENDOR CHQ 40218", "", "412.30", "15,780.40"),
        ("03/28/2026", "UTILITIES CONSOLIDATED", "318.64", "", "15,461.76"),
    ]
    y = 672
    for date, desc, wd, dep, bal in rows:
        _row(c, y, [(date, 60, "l"), (desc, 120, "l"), (wd, 400, "r"), (dep, 470, "r"), (bal, 545, "r")])
        y -= 16
    c.setFont(FONT + "-Bold", 9)
    _row(c, y - 10, [("Closing balance", 120, "l"), ("15,461.76", 545, "r")])
    c.save()
    return path


def statement_single_amount(path):
    """Date | Details | Amount | Balance, with signs carried by the balance."""
    c = canvas.Canvas(str(path), pagesize=letter)
    c.setFont(FONT + "-Bold", 11)
    c.drawString(60, 740, "SOUTHPORT SAVINGS - ACCOUNT SUMMARY")
    c.setFont(FONT + "-Bold", 9)
    _row(c, 700, [("Date", 60, "l"), ("Details", 130, "l"), ("Amount", 430, "r"), ("Balance", 520, "r")])
    c.setFont(FONT, 9)
    rows = [
        ("2026-04-01", "INTEREST CREDIT", "42.18", "8,042.18"),
        ("2026-04-03", "TRANSFER TO CHECKING", "(1,000.00)", "7,042.18"),
        ("2026-04-09", "DEPOSIT BRANCH 22", "2,500.00", "9,542.18"),
        ("2026-04-17", "SERVICE FEE", "(15.00)", "9,527.18"),
    ]
    y = 682
    for date, desc, amount, bal in rows:
        _row(c, y, [(date, 60, "l"), (desc, 130, "l"), (amount, 430, "r"), (bal, 520, "r")])
        y -= 16
    c.save()
    return path


def invoice(path, *, number="INV-2026-0142", vendor="Brightline Media Ltd"):
    c = canvas.Canvas(str(path), pagesize=letter)
    c.setFont(FONT + "-Bold", 16)
    c.drawString(60, 740, vendor)
    c.setFont(FONT, 9)
    c.drawString(60, 724, "18 Carrow Road, Manchester M1 4BT")
    c.setFont(FONT + "-Bold", 13)
    c.drawString(60, 690, "INVOICE")
    c.setFont(FONT, 10)
    c.drawString(60, 670, f"Invoice Number: {number}")
    c.drawString(60, 656, "Invoice Date: 12 May 2026")
    c.drawString(60, 642, "Due Date: 11 June 2026")
    c.drawString(60, 628, "Bill To: Coolrunner Plays Ltd")

    c.setFont(FONT + "-Bold", 10)
    _row(c, 596, [("Description", 60, "l"), ("Qty", 360, "r"), ("Unit", 430, "r"), ("Amount", 520, "r")])
    c.setFont(FONT, 10)
    items = [
        ("Creative campaign production", "1", "3,200.00", "3,200.00"),
        ("Media buying management fee", "1", "950.00", "950.00"),
        ("Asset licensing (12 months)", "4", "150.00", "600.00"),
    ]
    y = 578
    for desc, qty, unit, amount in items:
        _row(c, y, [(desc, 60, "l"), (qty, 360, "r"), (unit, 430, "r"), (amount, 520, "r")])
        y -= 16

    c.setFont(FONT, 10)
    _row(c, y - 14, [("Subtotal", 430, "r"), ("4,750.00", 520, "r")])
    _row(c, y - 30, [("VAT 20%", 430, "r"), ("950.00", 520, "r")])
    c.setFont(FONT + "-Bold", 11)
    _row(c, y - 48, [("Total Due", 430, "r"), ("5,700.00", 520, "r")])
    c.save()
    return path


if __name__ == "__main__":
    import sys
    out = sys.argv[1] if len(sys.argv) > 1 else "."
    statement_debit_credit(f"{out}/statement_dc.pdf")
    statement_single_amount(f"{out}/statement_amount.pdf")
    invoice(f"{out}/invoice_brightline.pdf")
    print("generated in", out)
