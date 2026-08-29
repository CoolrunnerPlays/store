import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import make_pdfs  # noqa: E402


@pytest.fixture(scope="session")
def docs(tmp_path_factory) -> dict[str, Path]:
    """Sample source documents, generated once per test session."""
    out = tmp_path_factory.mktemp("docs")
    return {
        "statement_dc": Path(make_pdfs.statement_debit_credit(out / "statement_dc.pdf")),
        "statement_amount": Path(make_pdfs.statement_single_amount(out / "statement_amount.pdf")),
        "invoice": Path(make_pdfs.invoice(out / "invoice.pdf")),
    }
