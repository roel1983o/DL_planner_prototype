from __future__ import annotations

import os
import uuid
import shutil
from pathlib import Path
from typing import Dict, Any

from fastapi import FastAPI, UploadFile, File, BackgroundTasks, HTTPException
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi import Request

from app.pipeline.pipeline import run_pipeline

APP_DIR = Path(__file__).resolve().parent
ASSETS_DIR = APP_DIR / "assets"
JOBS_DIR = Path(os.environ.get("JOBS_DIR", "/tmp/krantenplanner_jobs"))

app = FastAPI(title="Krantenplanner")

app.mount("/static", StaticFiles(directory=str(APP_DIR / "static")), name="static")
templates = Jinja2Templates(directory=str(APP_DIR / "templates"))

# In-memory status store (sufficient for a single Render web service instance)
JOBS: Dict[str, Dict[str, Any]] = {}

def _set(job_id: str, **kwargs: Any) -> None:
    JOBS.setdefault(job_id, {}).update(kwargs)

def _run_job(job_id: str, uploaded_path: str) -> None:
    try:
        _set(job_id, status="running", step="PARSER")
        workdir = JOBS_DIR / job_id
        out = run_pipeline(uploaded_parser_input=uploaded_path, workdir=str(workdir), assets_dir=str(ASSETS_DIR))
        _set(job_id, status="done", step="done", deel1_xlsx=out.deel1_xlsx, deel2_pdf=out.deel2_pdf)
    except Exception as e:
        _set(job_id, status="error", step="error", error=str(e))

@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/api/run")
async def run(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    if not file.filename:
        raise HTTPException(status_code=400, detail="Geen bestand ontvangen.")
    job_id = str(uuid.uuid4())
    workdir = JOBS_DIR / job_id
    workdir.mkdir(parents=True, exist_ok=True)

    uploaded_path = workdir / file.filename
    with uploaded_path.open("wb") as f:
        shutil.copyfileobj(file.file, f)

    _set(job_id, status="queued", step="queued")
    background_tasks.add_task(_run_job, job_id, str(uploaded_path))
    return {"job_id": job_id}

@app.get("/api/status/{job_id}")
def status(job_id: str):
    if job_id not in JOBS:
        raise HTTPException(status_code=404, detail="Onbekende job_id")
    return JOBS[job_id]

@app.get("/api/download/deel1/{job_id}")
def download_deel1(job_id: str):
    job = JOBS.get(job_id)
    if not job or job.get("status") != "done":
        raise HTTPException(status_code=400, detail="Output nog niet klaar.")
    path = job.get("deel1_xlsx")
    if not path or not os.path.exists(path):
        raise HTTPException(status_code=404, detail="DEEL 1 bestand niet gevonden.")
    return FileResponse(path, filename="Krantenplanning.xlsx")

@app.get("/api/download/deel2/{job_id}")
def download_deel2(job_id: str):
    job = JOBS.get(job_id)
    if not job or job.get("status") != "done":
        raise HTTPException(status_code=400, detail="Output nog niet klaar.")
    path = job.get("deel2_pdf")
    if not path or not os.path.exists(path):
        raise HTTPException(status_code=404, detail="DEEL 2 PDF niet gevonden.")
    return FileResponse(path, filename="Krantenplanning_handout.pdf")
