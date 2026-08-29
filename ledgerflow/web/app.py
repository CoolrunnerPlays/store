"""A small local web app: drop the workbook and the week's documents in, read
the audit of what was found, then approve the append.

Nothing is written to a spreadsheet until the audit has been shown and the
Approve button pressed. Everything stays on the machine it runs on; uploads go
to a per-job temporary folder and are not sent anywhere.
"""

from __future__ import annotations

import shutil
import tempfile
import uuid
from dataclasses import asdict
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Any

from flask import Flask, jsonify, request, send_file

from ..append import AppendResult, PreservationError, append_records, choose_sheet, plan_append
from ..extract import extract_records
from ..mapping import SYNONYMS, Profile, suggest_profile, value_for
from ..models import Flag, Record

MAX_UPLOAD_MB = 64

app = Flask(__name__, static_folder="static", template_folder="templates")
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024

JOBS: dict[str, dict[str, Any]] = {}
"""In-memory job store. This is a single-user local tool, so a dict is enough."""


def _json_safe(value: Any) -> Any:
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, date):
        return value.isoformat()
    return value


def _flag_json(flag: Flag) -> dict[str, Any]:
    return asdict(flag)


def _audit_rows(result: AppendResult, profile: Profile) -> list[dict[str, Any]]:
    """Every extracted record with its verdict and the cells it would fill.

    One list covering all four outcomes, so the browser can render the whole
    audit in a single table and colour each row by what will happen to it.
    """
    rows: list[dict[str, Any]] = []

    def add(record: Record, status: str, reason: str = "", matched_row: int | None = None) -> None:
        rows.append(
            {
                "status": status,
                "reason": reason,
                "matched_row": matched_row,
                "source": record.source,
                "kind": record.kind,
                "confidence": record.confidence,
                "cells": {
                    letter: _json_safe(value_for(record, field))
                    for letter, field in profile.columns.items()
                },
            }
        )

    for record in result.added:
        add(record, "new")
    for judgement in result.duplicates:
        add(judgement.record, "duplicate", judgement.match.reason, judgement.match.row)
    for judgement in result.probable_duplicates:
        add(judgement.record, "probable_duplicate", judgement.match.reason, judgement.match.row)
    for judgement in result.unmatchable:
        add(judgement.record, "unmatchable", judgement.match.reason)
    for record in result.skipped_other_kind:
        add(record, "other_kind", f"This sheet takes {' and '.join(profile.kinds)} rows.")
    return rows


def _audit_payload(result: AppendResult, profile: Profile) -> dict[str, Any]:
    return {
        "status": result.status,
        "message": result.message,
        "rows": _audit_rows(result, profile),
        "counts": {
            "new": len(result.added),
            "duplicate": len(result.duplicates),
            "probable_duplicate": len(result.probable_duplicates),
            "unmatchable": len(result.unmatchable),
            "other_kind": len(result.skipped_other_kind),
            "flags": len(result.flags),
        },
        "first_new_row": result.first_new_row,
        "can_append": result.status in ("written", "dry_run") and bool(result.added),
        "flags": [_flag_json(f) for f in result.flags],
    }


@app.get("/")
def index():
    return (Path(__file__).parent / "templates" / "index.html").read_text()


@app.get("/api/fields")
def fields():
    """The record fields a column can be mapped to, for the mapping dropdowns."""
    return jsonify(sorted(SYNONYMS.keys()))


@app.post("/api/analyze")
def analyze():
    """Read everything, detect the layout, and audit it. Writes nothing."""
    workbook = request.files.get("workbook")
    if workbook is None or not workbook.filename:
        return jsonify(error="Choose the Excel template you keep your records in."), 400
    if not workbook.filename.lower().endswith((".xlsx", ".xlsm")):
        return jsonify(error="The template needs to be an .xlsx or .xlsm file."), 400

    job_id = uuid.uuid4().hex[:12]
    workdir = Path(tempfile.mkdtemp(prefix=f"ledgerflow-{job_id}-"))
    book_path = workdir / Path(workbook.filename).name
    workbook.save(book_path)

    documents: list[Path] = []
    for upload in request.files.getlist("documents"):
        if not upload.filename:
            continue
        target = workdir / Path(upload.filename).name
        upload.save(target)
        documents.append(target)

    if not documents:
        shutil.rmtree(workdir, ignore_errors=True)
        return jsonify(error="Add at least one statement or invoice."), 400

    try:
        default_sheet, tables = choose_sheet(book_path)
    except ValueError as error:
        shutil.rmtree(workdir, ignore_errors=True)
        return jsonify(error=str(error)), 400

    records: list[Record] = []
    flags: list[Flag] = []
    per_file = []
    for path in documents:
        found, issues = extract_records(path)
        records.extend(found)
        flags.extend(issues)
        per_file.append(
            {
                "name": path.name,
                "records": len(found),
                "kind": sorted({r.kind for r in found}) or ["unreadable"],
                "flags": len(issues),
            }
        )

    profile = suggest_profile(tables[default_sheet])
    JOBS[job_id] = {"dir": workdir, "book": book_path, "records": records, "flags": flags, "tables": tables}

    audit = plan_append(book_path, records, profile, flags=flags)

    return jsonify(
        job=job_id,
        workbook=book_path.name,
        sheets={
            name: {
                "header_row": t.header_row,
                "first_data_row": t.first_data_row,
                "last_data_row": t.last_data_row,
                "rows": t.row_count,
                "totals_rows": t.total_rows,
                "columns": [
                    {
                        "letter": c.letter,
                        "header": c.header,
                        "type": c.inferred_type,
                        "computed": c.is_computed,
                        "formula": c.formula_template,
                    }
                    for c in t.columns
                ],
            }
            for name, t in tables.items()
        },
        default_sheet=default_sheet,
        profile=asdict(profile),
        files=per_file,
        audit=_audit_payload(audit, profile),
    )


@app.post("/api/preview")
def preview():
    """Re-run the audit after the mapping has been edited in the browser."""
    payload = request.get_json(force=True)
    job = JOBS.get(payload.get("job", ""))
    if job is None:
        return jsonify(error="That session has expired. Upload the files again."), 404

    profile = Profile(**payload["profile"])
    audit = plan_append(job["book"], job["records"], profile, flags=job["flags"])
    return jsonify(_audit_payload(audit, profile))


@app.post("/api/commit")
def commit():
    """Write the file, but only when the audit found something new to write."""
    payload = request.get_json(force=True)
    job = JOBS.get(payload.get("job", ""))
    if job is None:
        return jsonify(error="That session has expired. Upload the files again."), 404

    profile = Profile(**payload["profile"])
    book: Path = job["book"]
    output = job["dir"] / f"{book.stem}_updated{book.suffix}"

    try:
        result = append_records(book, job["records"], profile, output_path=output, flags=job["flags"])
    except PreservationError as error:
        return jsonify(error=str(error), status="preservation_failed"), 500

    if not result.file_written:
        # The safety check fired: the workbook was never opened for modification.
        return jsonify(
            status=result.status,
            message=result.message,
            download=None,
            audit=_audit_payload(result, profile),
        )

    job["output"] = result.output

    return jsonify(
        status="written",
        download=f"/download/{payload['job']}",
        filename=output.name,
        verified=result.verified,
        summary={
            "invoices": len(result.invoices),
            "transactions": len(result.transactions),
            "money_in": float(result.money_in),
            "money_out": float(result.money_out),
            "invoice_total": float(result.invoice_total),
            "net_added": float(result.total_added),
            "duplicates": len(result.duplicates),
            "probable_duplicates": len(result.probable_duplicates),
            "unmatchable": len(result.unmatchable),
            "other_kind": len(result.skipped_other_kind),
            "flags": len(result.flags),
            "formulas_updated": result.formulas_updated,
            "first_new_row": result.first_new_row,
            "sheet": result.sheet,
        },
        flags=[_flag_json(f) for f in result.flags],
    )


@app.get("/download/<job_id>")
def download(job_id: str):
    job = JOBS.get(job_id)
    if job is None or not job.get("output"):
        return jsonify(error="Nothing to download for that session."), 404
    return send_file(job["output"], as_attachment=True, download_name=Path(job["output"]).name)


def run(host: str = "127.0.0.1", port: int = 5000, debug: bool = False) -> None:
    print(f"\n  LedgerFlow is running at http://{host}:{port}\n")
    app.run(host=host, port=port, debug=debug)
