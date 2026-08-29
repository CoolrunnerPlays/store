import pytest
from openpyxl import load_workbook

from fixtures import build_ledger, build_populated_ledger
from ledgerflow.append import NO_NEW_MESSAGE
from ledgerflow.extract import extract_records
from ledgerflow.web.app import app


@pytest.fixture
def client():
    app.config["TESTING"] = True
    return app.test_client()


def upload(client, book, documents):
    return client.post(
        "/api/analyze",
        data={
            "workbook": (open(book, "rb"), book.name),
            "documents": [(open(d, "rb"), d.name) for d in documents],
        },
        content_type="multipart/form-data",
    )


@pytest.fixture
def analyzed(client, tmp_path, docs):
    book = build_ledger(tmp_path / "book.xlsx")
    response = upload(client, book, [docs["statement_dc"], docs["invoice"]])
    assert response.status_code == 200
    return response.get_json()


def test_analyze_reports_the_layout_and_the_audit(analyzed):
    assert analyzed["default_sheet"] == "Transactions"
    assert analyzed["sheets"]["Transactions"]["header_row"] == 3
    audit = analyzed["audit"]
    assert audit["counts"]["new"] == 8
    assert audit["counts"]["other_kind"] == 1
    assert audit["can_append"] is True


def test_every_extracted_record_appears_in_the_audit_with_a_verdict(analyzed):
    rows = analyzed["audit"]["rows"]
    assert len(rows) == 9
    assert {r["status"] for r in rows} == {"new", "other_kind"}
    assert all("cells" in r for r in rows)


def test_analyze_rejects_a_non_workbook(client, docs, tmp_path):
    response = client.post(
        "/api/analyze",
        data={"workbook": (open(docs["invoice"], "rb"), "invoice.pdf")},
        content_type="multipart/form-data",
    )
    assert response.status_code == 400
    assert "xlsx" in response.get_json()["error"]


def test_analyze_needs_at_least_one_document(client, tmp_path):
    book = build_ledger(tmp_path / "b.xlsx")
    response = client.post(
        "/api/analyze",
        data={"workbook": (open(book, "rb"), "b.xlsx")},
        content_type="multipart/form-data",
    )
    assert response.status_code == 400


def test_preview_reflects_an_edited_mapping(client, analyzed):
    profile = analyzed["profile"]
    del profile["columns"]["C"]
    body = client.post("/api/preview", json={"job": analyzed["job"], "profile": profile}).get_json()
    assert body["rows"] and "C" not in body["rows"][0]["cells"]


def test_commit_writes_a_file_and_reports_the_summary(client, analyzed, tmp_path):
    body = client.post("/api/commit", json={"job": analyzed["job"], "profile": analyzed["profile"]}).get_json()
    assert body["status"] == "written"
    assert body["verified"] is True
    assert body["summary"]["transactions"] == 8

    download = client.get(body["download"])
    assert download.status_code == 200
    out = tmp_path / "downloaded.xlsx"
    out.write_bytes(download.data)
    assert load_workbook(out)["Transactions"]["D21"].value == "=SUM(D4:D19)"


def test_a_fully_populated_upload_is_refused_at_the_audit(client, tmp_path, docs):
    """The audit scenario, end to end through the browser API."""
    records = extract_records(docs["statement_dc"])[0]
    book = build_populated_ledger(tmp_path / "full.xlsx", records, variant="textdates")

    analyzed = upload(client, book, [docs["statement_dc"]]).get_json()
    audit = analyzed["audit"]
    assert audit["status"] == "no_new_records"
    assert audit["message"] == NO_NEW_MESSAGE
    assert audit["counts"]["new"] == 0
    assert audit["counts"]["duplicate"] == len(records)
    assert audit["can_append"] is False
    assert all(r["status"] == "duplicate" for r in audit["rows"])


def test_commit_refuses_to_write_when_the_audit_found_nothing(client, tmp_path, docs):
    records = extract_records(docs["statement_dc"])[0]
    book = build_populated_ledger(tmp_path / "full.xlsx", records, variant="reworded")
    analyzed = upload(client, book, [docs["statement_dc"]]).get_json()

    body = client.post("/api/commit", json={"job": analyzed["job"], "profile": analyzed["profile"]}).get_json()
    assert body["status"] == "no_new_records"
    assert body["download"] is None
    assert client.get(f"/download/{analyzed['job']}").status_code == 404


def test_an_expired_session_says_so(client):
    response = client.post("/api/commit", json={"job": "nope", "profile": {"sheet": "x"}})
    assert response.status_code == 404
