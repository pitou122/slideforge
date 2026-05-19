# Project: Video → PPTX Web App

## Overview
Build a web app that takes a screen-recorded presentation video as input,
extracts slides automatically, and reconstructs a faithful PPTX file for download.
Read `pipeline_spec.md` for the full technical spec before writing any code.

---

## Tech Stack
- **Backend**: Python + FastAPI
- **Frontend**: Single HTML file (vanilla JS, no framework)
- **Key libraries**: opencv-python, pillow, scikit-image, anthropic, python-pptx, python-dotenv

## Project Structure
```
slide-tool/
├── CLAUDE.md               ← this file
├── pipeline_spec.md        ← full pipeline spec (READ THIS FIRST)
├── .env                    ← ANTHROPIC_API_KEY (never commit)
├── .gitignore
├── requirements.txt
├── main.py                 ← FastAPI entry point
├── src/
│   ├── extractor.py        ← Stage 1: frame extraction
│   ├── detector.py         ← Stage 2: slide change detection
│   ├── vision.py           ← Stage 3: Claude Vision API
│   ├── cropper.py          ← Stage 4: image cropping
│   └── builder.py          ← Stage 5: PPTX generation
└── frontend/
    └── index.html          ← Single-file UI
```

## Implementation Plan

### Step 1 — Scaffold
- Create all files and folders above
- Write `requirements.txt`
- Write `.gitignore` (exclude `.env`, `jobs/`, `__pycache__`)

### Step 2 — Backend: src/ modules
Implement each stage as a standalone module following `pipeline_spec.md` exactly.
Each module exposes one main function:
- `extractor.py` → `extract_frames(video_path, out_dir, fps) → list[FrameMeta]`
- `detector.py`  → `detect_slides(frames, out_dir, threshold) → list[SlideMeta]`
- `vision.py`    → `analyze_slide(slide, client) → dict`
- `cropper.py`   → `crop_images(extracted, out_dir) → None`
- `builder.py`   → `build_pptx(extracted, out_dir) → Path`

### Step 3 — Backend: main.py (FastAPI)
Three routes:
- `POST /upload` — accept video file + fps + threshold params, start background job, return job_id
- `GET  /status/{job_id}` — return { status, progress (0-100), message }
- `GET  /download/{job_id}` — stream the output.pptx file as download

Job state stored in a simple in-memory dict (no database needed).
Run pipeline stages sequentially in a background thread.
Update progress after each stage completes.

### Step 4 — Frontend: frontend/index.html
Single HTML file served by FastAPI StaticFiles.
UI flow:
1. Drag-and-drop or click-to-browse for video file
2. Optional sliders: FPS (default 1.0) and SSIM threshold (default 0.85)
3. Upload button → POST to /upload → receive job_id
4. Poll GET /status/{job_id} every 2 seconds → show progress bar + status message
5. When status = "done" → show Download button → GET /download/{job_id}
6. When status = "error" → show error message in red

Design: dark background (#0a0a0a), monospace font, orange accent (#f97316).
Keep it minimal and functional — no CSS frameworks.

### Step 5 — Environment
- Load ANTHROPIC_API_KEY from .env via python-dotenv
- Never hardcode the key anywhere in source files
- Print a clear error on startup if key is missing

## Run Command
```bash
uvicorn main:app --reload --port 8000
```
Then open http://localhost:8000

## Key Constraints
- Do NOT use Node.js or pptxgenjs — use python-pptx only
- Do NOT use any database — in-memory dict is enough
- The frontend must be a single index.html file (inline CSS + JS)
- API key comes exclusively from .env file
- Follow pipeline_spec.md for all business logic — do not invent your own logic
