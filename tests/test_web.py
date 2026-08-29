import pytest
from openpyxl import load_workbook

from fixtures import build_ledger
from ledgerflow.web.app import app


@pytest.fixture
def client():
    app.config["TESTING"] = True
    return app.test_client()


@pytest.fixture
def analyzed(client, tmp_path, docs):
    book = build_ledger(tmp_path / "book.xlsx")
    payload = {
        "workbook": (open(book, "rb"), "book.xlsx"),
        "documents": [
            (open(docs["statement_dc"], "rb"), "statement_dc.pdf"),
            (open(docs["invoice"], "rb"), "invoice.pdf"),
        ],
    }
    response = client.post("/api/analyze", data=payload, content_type="multipart/form-data")
    assert response.status_code == 200
    return response.get_json()


def test_analyze_reports_the_layout_and_what_was_found(analyzed):
    assert analyzed["default_sheet"] == "Transactions"
    assert analyzed["sheets"]["Transactions"]["header_row"] == 3
    assert {f["name"] for f in analyzed["files"]} == {"statement_dc.pdf", "invoice.pdf"}
    assert len(analyzed["preview"]) == 8
    assert len(analyzed["other_kind"]) == 1


def test_analyze_rejects_a_non_workbook(client, docs):
    response = client.post(
        "/api/analyze",
        data={"workbook": (open(docs["invoice"], "rb"), "invoice.pdf")},
        content_type="multipart/form-data",
    )
    assert response.status_code == 400
    assert "xlsx" in response.get_json()["error"]


def test_preview_reflects_an_edited_mapping(client, analyzed):
    profile = analyzed["profile"]
    del profile["columns"]["C"]
    response = client.post("/api/preview", json={"job": analyzed["job"], "profile": profile})
    rows = response.get_json()["rows"]
    assert rows and "C" not in rows[0]["cells"]


def test_commit_writes_a_file_and_reports_the_summary(client, analyzed, tmp_path):
    response = client.post("/api/commit", json={"job": analyzed["job"], "profile": analyzed["profile"]})
    body = response.get_json()
    assert body["summary"]["transactions"] == 8
    assert body["summary"]["formulas_updated"] > 0

    download = client.get(body["download"])
    assert download.status_code == 200
    out = tmp_path / "downloaded.xlsx"
    out.write_bytes(download.data)
    ws = load_workbook(out)["Transactions"]
    assert ws["D21"].value == "=SUM(D4:D19)"


def test_an_expired_session_says_so(client):
    response = client.post("/api/commit", json={"job": "nope", "profile": {"sheet": "x"}})
    assert response.status_code == 404
