from __future__ import annotations

import json
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, Response

from backend.core import analyze_records, parse_fit_records
from backend.excel_export import build_excel

app = FastAPI(title="Burst Analyzer API", version="1.1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

PROJECT_ROOT = Path(__file__).resolve().parent.parent
HTML_FILE = PROJECT_ROOT / "burst_analyzer.html"


@app.get("/")
async def index() -> FileResponse:
    if not HTML_FILE.exists():
        raise HTTPException(status_code=404, detail="burst_analyzer.html non trovato")
    return FileResponse(HTML_FILE)


@app.get("/api/health")
async def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/api/analyze")
async def analyze_fit(
    file: UploadFile = File(...),
    thresholds_json: str = Form(...),
    min_dur: int = Form(4),
    merge_gap: int = Form(2),
) -> dict:
    if not file.filename or not file.filename.lower().endswith(".fit"):
        raise HTTPException(status_code=400, detail="File must be .fit")

    try:
        thresholds = json.loads(thresholds_json)
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail="thresholds_json non valido") from exc

    if not isinstance(thresholds, list) or not thresholds:
        raise HTTPException(status_code=400, detail="Serve almeno una soglia")

    try:
        file_bytes = await file.read()
        records = parse_fit_records(file_bytes)
        all_results = analyze_records(records, thresholds, int(min_dur), int(merge_gap))
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Errore analisi server: {exc}") from exc

    duration_sec = records[-1]["time_sec"] if records else 0

    return {
        "records": records,
        "allResults": all_results,
        "stats": {
            "samples": len(records),
            "durationSec": duration_sec,
            "fileName": file.filename,
        },
    }


@app.post("/api/export-excel")
async def export_excel(
    results_json: str = Form(...),
    min_dur: int = Form(4),
    filename: str = Form("burst_analysis"),
) -> Response:
    try:
        all_results = json.loads(results_json)
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail="results_json non valido") from exc

    try:
        xlsx_bytes = build_excel(all_results, min_dur=int(min_dur))
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Errore generazione Excel: {exc}") from exc

    safe_name = "".join(c for c in filename if c.isalnum() or c in "-_ .")
    safe_name = " ".join(safe_name.split())
    if not safe_name:
        safe_name = "burst_analysis"
    if not safe_name.endswith(".xlsx"):
        safe_name += ".xlsx"

    return Response(
        content=xlsx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{safe_name}"'},
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="127.0.0.1", port=8010, reload=False, workers=1)