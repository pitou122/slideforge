"""FastAPI entry point for the Video → PPTX web app.

The pipeline is driven stage by stage from the UI:

    upload → extract & detect → [select slides] → vision → crop → build

Each stage is its own endpoint, runs in a background thread, and persists
its state to ``jobs/{job_id}/job.json`` so progress survives a server
restart.
"""
from __future__ import annotations

import json
import os
import sys
import threading
import uuid
from pathlib import Path

from dotenv import load_dotenv
from fastapi import Body, FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles

from src.extractor import extract_slides
from src.vision import analyze_slide
from src.cropper import crop_images
from src.builder import build_pptx

# --- Environment -----------------------------------------------------------
load_dotenv()
API_KEY = os.getenv("ANTHROPIC_API_KEY")
if not API_KEY:
    sys.stderr.write(
        "\nERROR: ANTHROPIC_API_KEY is not set.\n"
        "Create a .env file in the project root containing:\n"
        "    ANTHROPIC_API_KEY=sk-ant-...\n\n"
    )
    sys.exit(1)

BASE_DIR = Path(__file__).parent
JOBS_DIR = BASE_DIR / "jobs"
JOBS_DIR.mkdir(exist_ok=True)

app = FastAPI(title="Video → PPTX")

# In-memory write-through cache of job state; the source of truth is job.json.
_jobs: dict[str, dict] = {}
_lock = threading.Lock()

# Stage that must be complete before a given stage may run.
PREREQ = {
    "extract": {"uploaded", "detected", "error"},
    "vision": {"detected", "analyzed", "error"},
    "crop": {"analyzed", "cropped", "error"},
    "build": {"cropped", "done", "error"},
}


# --- Job state -------------------------------------------------------------
def _job_dir(job_id: str) -> Path:
    return JOBS_DIR / job_id


def _job_file(job_id: str) -> Path:
    return _job_dir(job_id) / "job.json"


def _load_job(job_id: str) -> dict | None:
    """Return job state from the cache, falling back to job.json on disk."""
    with _lock:
        if job_id in _jobs:
            return dict(_jobs[job_id])
    jf = _job_file(job_id)
    if jf.exists():
        with open(jf, encoding="utf-8") as fh:
            job = json.load(fh)
        with _lock:
            _jobs[job_id] = job
        return dict(job)
    return None


def _save_job(job: dict) -> None:
    """Persist job state to both the cache and job.json."""
    with _lock:
        _jobs[job["job_id"]] = dict(job)
    with open(_job_file(job["job_id"]), "w", encoding="utf-8") as fh:
        json.dump(job, fh, indent=2)


def _update(job_id: str, **fields) -> None:
    """Apply ``fields`` to a job and persist it."""
    job = _load_job(job_id)
    if job is None:
        return
    job.update(fields)
    _save_job(job)


def _require(job_id: str, stage: str) -> dict:
    """Fetch a job and verify its prerequisite stage is complete."""
    job = _load_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="Unknown job_id")
    if job["status"] not in PREREQ[stage]:
        raise HTTPException(
            status_code=409,
            detail=f"Cannot run '{stage}' while job status is '{job['status']}'",
        )
    return job


def _run_bg(target, *args) -> None:
    """Start ``target`` in a daemon background thread."""
    threading.Thread(target=target, args=args, daemon=True).start()


# --- Stage workers ---------------------------------------------------------
def _do_extract(job_id: str) -> None:
    job = _load_job(job_id)
    out_dir = _job_dir(job_id)
    try:
        _update(job_id, status="detecting", progress=0,
                message="Scanning video for slide changes…")

        def on_progress(pct: int, msg: str) -> None:
            _update(job_id, progress=pct, message=msg)

        slides = extract_slides(
            out_dir / "input_video.mp4", out_dir, job["interval"],
            progress_cb=on_progress,
        )
        _update(job_id, status="detected", progress=100,
                message=f"Detected {len(slides)} slide(s)",
                selected=[s["slide_id"] for s in slides])
    except Exception as exc:  # noqa: BLE001
        _update(job_id, status="error", message=f"{type(exc).__name__}: {exc}")


def _do_vision(job_id: str) -> None:
    out_dir = _job_dir(job_id)
    try:
        job = _load_job(job_id)
        from anthropic import Anthropic
        client = Anthropic(api_key=API_KEY)

        with open(out_dir / "slides_detected.json", encoding="utf-8") as fh:
            detected = json.load(fh)
        selected = set(job.get("selected") or [s["slide_id"] for s in detected])
        slides = [s for s in detected if s["slide_id"] in selected]

        _update(job_id, status="analyzing", progress=0,
                message=f"Analyzing {len(slides)} slide(s) with Claude Vision…")

        results: list[dict] = []
        for i, slide in enumerate(slides):
            resolved = dict(slide)
            resolved["representative_frame"] = str(
                out_dir / slide["representative_frame"]
            )
            try:
                results.append(analyze_slide(resolved, client))
            except Exception as exc:  # noqa: BLE001
                print(f"[vision] ERROR on {slide['slide_id']}: {exc}")
                results.append(
                    {"slide_id": slide["slide_id"], "error": True, "columns": []}
                )
            _update(job_id, progress=int((i + 1) / len(slides) * 100),
                    message=f"Analyzed {i + 1}/{len(slides)} slide(s)")

        with open(out_dir / "slides_extracted.json", "w", encoding="utf-8") as fh:
            json.dump(results, fh, indent=2)

        errs = sum(1 for r in results if r.get("error"))
        _update(job_id, status="analyzed", progress=100,
                message=f"Vision complete ({len(results)} slides, {errs} error(s))")
    except Exception as exc:  # noqa: BLE001
        _update(job_id, status="error", message=f"{type(exc).__name__}: {exc}")


def _do_crop(job_id: str) -> None:
    out_dir = _job_dir(job_id)
    try:
        _update(job_id, status="cropping", progress=10,
                message="Cropping charts and images…")
        with open(out_dir / "slides_extracted.json", encoding="utf-8") as fh:
            extracted = json.load(fh)
        crop_images(extracted, out_dir)
        n = len(list((out_dir / "crops").glob("*.png"))) if (out_dir / "crops").exists() else 0
        _update(job_id, status="cropped", progress=100,
                message=f"Cropped {n} image(s)")
    except Exception as exc:  # noqa: BLE001
        _update(job_id, status="error", message=f"{type(exc).__name__}: {exc}")


def _do_build(job_id: str) -> None:
    out_dir = _job_dir(job_id)
    try:
        _update(job_id, status="building", progress=20,
                message="Building PPTX…")
        with open(out_dir / "slides_extracted.json", encoding="utf-8") as fh:
            extracted = json.load(fh)
        build_pptx(extracted, out_dir)
        _update(job_id, status="done", progress=100,
                message="Done — your presentation is ready")
    except Exception as exc:  # noqa: BLE001
        _update(job_id, status="error", message=f"{type(exc).__name__}: {exc}")


# --- Routes ----------------------------------------------------------------
@app.post("/upload")
async def upload(video: UploadFile = File(...), interval: float = Form(5.0)):
    """Accept a video, create a job, and return its job_id."""
    job_id = uuid.uuid4().hex[:12]
    job_dir = _job_dir(job_id)
    job_dir.mkdir(parents=True, exist_ok=True)

    with open(job_dir / "input_video.mp4", "wb") as fh:
        while chunk := await video.read(1024 * 1024):
            fh.write(chunk)

    _save_job({
        "job_id": job_id,
        "status": "uploaded",
        "progress": 0,
        "message": "Uploaded — ready to extract",
        "interval": interval,
        "selected": [],
    })
    return {"job_id": job_id}


@app.post("/jobs/{job_id}/extract")
async def run_extract(job_id: str):
    """Run the merged Extract & Detect stage in the background."""
    _require(job_id, "extract")
    _run_bg(_do_extract, job_id)
    return {"started": "extract"}


@app.get("/jobs/{job_id}/slides")
async def get_slides(job_id: str):
    """Return the detected slides with image URLs and current selection."""
    job = _load_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="Unknown job_id")
    path = _job_dir(job_id) / "slides_detected.json"
    if not path.exists():
        return {"slides": []}
    with open(path, encoding="utf-8") as fh:
        detected = json.load(fh)
    selected = set(job.get("selected") or [s["slide_id"] for s in detected])
    return {
        "slides": [
            {
                "slide_id": s["slide_id"],
                "start_sec": s["start_sec"],
                "end_sec": s["end_sec"],
                "image_url": f"/jobs/{job_id}/image?rel={s['representative_frame']}",
                "selected": s["slide_id"] in selected,
            }
            for s in detected
        ]
    }


@app.post("/jobs/{job_id}/select")
async def select_slides(job_id: str, slide_ids: list[str] = Body(..., embed=True)):
    """Persist which detected slides should continue past detection."""
    job = _load_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="Unknown job_id")
    _update(job_id, selected=slide_ids)
    return {"selected": slide_ids}


@app.post("/jobs/{job_id}/vision")
async def run_vision(job_id: str):
    """Run Claude Vision analysis on the selected slides."""
    job = _require(job_id, "vision")
    if not job.get("selected"):
        raise HTTPException(status_code=400, detail="No slides selected")
    _run_bg(_do_vision, job_id)
    return {"started": "vision"}


@app.post("/jobs/{job_id}/crop")
async def run_crop(job_id: str):
    """Run the image-cropping stage."""
    _require(job_id, "crop")
    _run_bg(_do_crop, job_id)
    return {"started": "crop"}


@app.post("/jobs/{job_id}/build")
async def run_build(job_id: str):
    """Run the PPTX-building stage."""
    _require(job_id, "build")
    _run_bg(_do_build, job_id)
    return {"started": "build"}


@app.get("/jobs/{job_id}/extracted")
async def get_extracted(job_id: str):
    """Return a per-slide summary of the Vision extraction."""
    path = _job_dir(job_id) / "slides_extracted.json"
    if not path.exists():
        return {"slides": []}
    with open(path, encoding="utf-8") as fh:
        data = json.load(fh)
    summary = []
    for s in data:
        cols = s.get("columns", []) or []
        summary.append({
            "slide_id": s.get("slide_id"),
            "layout": s.get("layout", "—"),
            "columns": len(cols),
            "charts": sum(
                1 for c in cols for e in c.get("elements", [])
                if e.get("type") == "image_crop"
            ),
            "error": bool(s.get("error")),
        })
    return {"slides": summary}


@app.get("/jobs/{job_id}/crops")
async def get_crops(job_id: str):
    """Return image URLs for every crop produced by the cropping stage."""
    crops_dir = _job_dir(job_id) / "crops"
    if not crops_dir.exists():
        return {"crops": []}
    names = sorted(p.name for p in crops_dir.glob("*.png"))
    return {
        "crops": [
            {"name": n, "image_url": f"/jobs/{job_id}/image?rel=crops/{n}"}
            for n in names
        ]
    }


@app.get("/jobs/{job_id}/image")
async def get_image(job_id: str, rel: str):
    """Serve an image from inside a job directory (path-restricted)."""
    job_dir = _job_dir(job_id).resolve()
    target = (job_dir / rel).resolve()
    if job_dir != target and job_dir not in target.parents:
        raise HTTPException(status_code=403, detail="Path outside job directory")
    if not target.is_file():
        raise HTTPException(status_code=404, detail="Image not found")
    return FileResponse(target)


@app.get("/jobs/{job_id}/status")
async def status(job_id: str):
    """Return the current status, progress (0-100), and message for a job."""
    job = _load_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="Unknown job_id")
    return {
        "status": job["status"],
        "progress": job["progress"],
        "message": job["message"],
    }


@app.get("/download/{job_id}")
async def download(job_id: str):
    """Stream the finished output.pptx as a file download."""
    job = _load_job(job_id)
    if job is None:
        raise HTTPException(status_code=404, detail="Unknown job_id")
    if job["status"] != "done":
        raise HTTPException(status_code=409, detail="Job is not finished")
    pptx_path = _job_dir(job_id) / "output.pptx"
    if not pptx_path.exists():
        raise HTTPException(status_code=404, detail="Output file missing")
    return FileResponse(
        pptx_path,
        media_type=(
            "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        ),
        filename="presentation.pptx",
    )


@app.get("/")
async def index():
    """Redirect the root URL to the static frontend."""
    return RedirectResponse(url="/static/index.html")


# Serve the single-file frontend.
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "frontend")), name="static")
