"""Stage 1 — Extract slides by fixed-interval capture.

Sample one frame every ``interval_sec`` seconds from the input video.
Each sampled frame becomes a slide candidate shown to the user for selection.
"""
from __future__ import annotations

import json
from pathlib import Path

import cv2


def extract_slides(
    video_path: str | Path,
    out_dir: str | Path,
    interval_sec: float,
    progress_cb=None,
) -> list[dict]:
    """Capture one frame every ``interval_sec`` seconds from ``video_path``.

    Each captured frame is saved as ``{out_dir}/slides/slide_NNN.png``.
    ``progress_cb`` is an optional ``callable(pct, message)`` for progress.

    Returns a list of SlideMeta dicts:
    ``{slide_id, representative_frame, start_sec, end_sec}``.
    """
    video_path = Path(video_path)
    out_dir = Path(out_dir)
    slides_dir = out_dir / "slides"
    slides_dir.mkdir(parents=True, exist_ok=True)

    cap = cv2.VideoCapture(str(video_path))
    if not cap.isOpened():
        raise RuntimeError(f"Could not open video: {video_path}")

    video_fps = cap.get(cv2.CAP_PROP_FPS) or 0.0
    if video_fps <= 0:
        video_fps = 30.0
    total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0)

    probe_step = max(1, int(video_fps * interval_sec))

    slides: list[dict] = []
    frame_index = 0

    while True:
        ok, frame = cap.read()
        if not ok:
            break

        if frame_index % probe_step == 0:
            ts = round(frame_index / video_fps, 3)
            slide_num = len(slides) + 1
            slide_id = f"slide_{slide_num:03d}"
            rel_path = f"slides/{slide_id}.png"
            cv2.imwrite(str(out_dir / rel_path), frame)
            slides.append({
                "slide_id": slide_id,
                "representative_frame": rel_path,
                "start_sec": ts,
                "end_sec": ts,
            })

            if progress_cb and total_frames:
                pct = int(frame_index / total_frames * 100)
                progress_cb(pct, f"Captured {slide_num} frame(s)…")

        frame_index += 1

    cap.release()

    with open(out_dir / "slides_detected.json", "w", encoding="utf-8") as fh:
        json.dump(slides, fh, indent=2)

    return slides
