"""A small local web app: drop the workbook and the week's documents in, check
what was found, then download the updated file.

Everything stays on the machine it runs on. Uploads go to a per-job temporary
folder and nothing is sent anywhere.
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

from ..append import append_records, choose_sheet
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


def _record_json(record: Record) -> dict[str, Any]:
    data = {k: _json_safe(v) for k, v in asdict(record).items()}
    data["fingerprint"] = record.fingerprint()
    return data


def _flag_json(flag: Flag) -> dict[str, Any]:
    return asdict(flag)


@app.get("/")
def index():
    return (Path(__file__).parent / "templates" / "index.html").read_text()


@app.get("/api/fields")
def fields():
    """The record fields a column can be mapped to, for the mapping dropdowns."""
    return jsonify(sorted(SYNONYMS.keys()))


@app.post("/api/analyze")
def analyze():
    """Read everything, detect the layout, and report without writing anything."""
    workbook = request.files.get("workbook")
    if workbook is None or not workbook.filename:
        return jsonify(error="Choose the spreadsheet you keep your records in."), 400
    if not workbook.filename.lower().endswith((".xlsx", ".xlsm")):
        return jsonify(error="The workbook needs to be an .xlsx or .xlsm file."), 400

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
    JOBS[job_id] = {
        "dir": workdir,
        "book": book_path,
        "records": records,
        "flags": flags,
        "tables": tables,
    }

    preview = append_records(book_path, records, profile, flags=flags, dry_run=True)

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
        preview=[_record_json(r) for r in preview.added],
        duplicates=[_record_json(r) for r in preview.skipped_duplicates],
        other_kind=[_record_json(r) for r in preview.skipped_other_kind],
        flags=[_flag_json(f) for f in flags],
    )


@app.post("/api/preview")
def preview():
    """Re-run the dry run after the mapping has been edited in the browser."""
    payload = request.get_json(force=True)
    job = JOBS.get(payload.get("job", ""))
    if job is None:
        return jsonify(error="That session has expired. Upload the files again."), 404

    profile = Profile(**payload["profile"])
    result = append_records(job["book"], job["records"], profile, flags=job["flags"], dry_run=True)
    rows = [
        {
            "cells": {
                letter: _json_safe(value_for(record, field))
                for letter, field in profile.columns.items()
            },
            "kind": record.kind,
            "source": record.source,
            "confidence": record.confidence,
        }
        for record in result.added
    ]
    return jsonify(
        rows=rows,
        added=len(result.added),
        duplicates=len(result.skipped_duplicates),
        other_kind=len(result.skipped_other_kind),
        first_new_row=result.first_new_row,
    )


@app.post("/api/commit")
def commit():
    """Write the file for real and hand back a download link."""
    payload = request.get_json(force=True)
    job = JOBS.get(payload.get("job", ""))
    if job is None:
        return jsonify(error="That session has expired. Upload the files again."), 404

    profile = Profile(**payload["profile"])
    book: Path = job["book"]
    output = job["dir"] / f"{book.stem}_updated{book.suffix}"

    result = append_records(book, job["records"], profile, output_path=output, flags=job["flags"])
    job["output"] = result.output if result.output.exists() else None

    return jsonify(
        download=f"/download/{payload['job']}" if job["output"] else None,
        filename=output.name,
        summary={
            "invoices": len(result.invoices),
            "transactions": len(result.transactions),
            "money_in": float(result.money_in),
            "money_out": float(result.money_out),
            "invoice_total": float(result.invoice_total),
            "net_added": float(result.total_added),
            "duplicates": len(result.skipped_duplicates),
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
