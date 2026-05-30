"""
main.py — FastAPI application for Thesis Auto-Corrector
"""

import os
import uuid
import json
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, JSONResponse
import aiofiles

# Add parent to path for imports
import sys
sys.path.insert(0, str(Path(__file__).parent))

from corrector.thesis_corrector import ThesisCorrector

app = FastAPI(
    title="Thesis Auto-Corrector",
    description="Автоматическая нормоконтроль дипломных работ по ПР V-08-2022",
    version="1.0.0"
)

RULES_PATH = Path(__file__).parent.parent / "data" / "norm_control_rules.json"
UPLOAD_DIR = Path("/tmp/thesis_uploads")
OUTPUT_DIR = Path("/tmp/thesis_outputs")
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


@app.get("/")
async def root():
    return {
        "service": "Thesis Auto-Corrector",
        "standard": "ПР V-08-2022 КарТУ",
        "endpoints": {
            "POST /correct": "Загрузить .docx, получить исправленный файл + отчёт",
            "GET /rules": "Просмотреть текущие правила нормоконтроля",
            "GET /health": "Статус сервиса"
        }
    }


@app.get("/health")
async def health():
    rules_exist = RULES_PATH.exists()
    return {"status": "ok", "rules_loaded": rules_exist}


@app.get("/rules")
async def get_rules():
    """Return the current norm control rules JSON."""
    if not RULES_PATH.exists():
        raise HTTPException(status_code=404, detail="Rules file not found")
    with open(RULES_PATH, encoding="utf-8") as f:
        return json.load(f)


@app.post("/correct")
async def correct_thesis(
    file: UploadFile = File(..., description="Файл дипломной работы (.docx)")
):
    """
    Upload a .docx thesis file.
    Returns: corrected .docx + text correction report.
    """
    if not file.filename.endswith(".docx"):
        raise HTTPException(status_code=400, detail="Только файлы .docx принимаются")

    job_id = str(uuid.uuid4())[:8]
    input_path = UPLOAD_DIR / f"{job_id}_input.docx"
    output_path = OUTPUT_DIR / f"{job_id}_corrected.docx"
    report_path = OUTPUT_DIR / f"{job_id}_report.txt"

    # Save upload
    async with aiofiles.open(input_path, "wb") as f:
        content = await file.read()
        await f.write(content)

    # Run corrector
    try:
        corrector = ThesisCorrector(str(RULES_PATH))
        report = corrector.correct(str(input_path), str(output_path))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка обработки: {str(e)}")

    # Save report text
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(report.summary())

    return JSONResponse({
        "job_id": job_id,
        "violations_total": len(report.violations),
        "auto_fixed": report.auto_fixed_count,
        "manual_review": report.manual_review_count,
        "download_corrected": f"/download/{job_id}/docx",
        "download_report": f"/download/{job_id}/report",
        "violations": [
            {
                "location": v.location,
                "rule": v.rule_ref,
                "description": v.description,
                "auto_fixed": v.auto_fixed,
                "original": v.original[:100] if v.original else None,
                "corrected": v.corrected
            }
            for v in report.violations
        ]
    })


@app.get("/download/{job_id}/docx")
async def download_corrected(job_id: str):
    path = OUTPUT_DIR / f"{job_id}_corrected.docx"
    if not path.exists():
        raise HTTPException(status_code=404, detail="Файл не найден")
    return FileResponse(
        path=str(path),
        filename=f"thesis_corrected_{job_id}.docx",
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


@app.get("/download/{job_id}/report")
async def download_report(job_id: str):
    path = OUTPUT_DIR / f"{job_id}_report.txt"
    if not path.exists():
        raise HTTPException(status_code=404, detail="Отчёт не найден")
    return FileResponse(
        path=str(path),
        filename=f"norm_control_report_{job_id}.txt",
        media_type="text/plain; charset=utf-8"
    )
