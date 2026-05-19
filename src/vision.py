"""Stage 3 — Claude Vision Extraction.

Send each representative slide image to the Claude Vision API and extract a
strict JSON description of its content. See pipeline_spec.md for the spec.
"""
from __future__ import annotations

import base64
import json
import time
from pathlib import Path

from PIL import Image

MODEL = "claude-sonnet-4-6"
MAX_TOKENS = 4000

SYSTEM_PROMPT = """You are an IELTS slide content extractor. Analyze this presentation slide image
and return a single strict JSON object describing every element needed to
faithfully reconstruct it in PowerPoint.

CONTEXT:
- White/light background with black or dark body text
- Title is large, often orange/amber colored, may be multi-line
- Layout is 1–3 columns of IELTS task examples
- Each column may contain: date label, task description paragraph,
  chart/graph/table image, highlighted task prompt box, TASK N badge
- Some text may be partially hidden — infer when needed

INCLUDE:
✓ Title (exact text, including numbering like "1/")
✓ Date labels (e.g. "13/5/2017")
✓ Task description paragraphs (infer if partially hidden)
✓ Highlighted/colored task prompt boxes
✓ Charts, graphs, tables → mark as image_crop with bounding box
✓ TASK N badges

EXCLUDE:
✗ Human faces, webcam overlays in corners
✗ Presentation UI chrome (cursor menus, toolbars, pointer icons)
✗ Any image clearly unrelated to IELTS academic content

SLIDE CONTENT AREA (report this first):
Identify the actual presentation slide canvas — the white/light rectangle that is the
slide itself — separate from any surrounding application chrome (ribbon, toolbar,
slide-thumbnail panel, title bar, taskbar, webcam overlay).
Report "slide_area": [x1, y1, x2, y2] as 0.0–1.0 fractions of the FULL input image.
If the full image IS the slide with no chrome, report [0.0, 0.0, 1.0, 1.0].

BOUNDING BOXES — required on EVERY element (text, badge, image — all types):
- bbox: [x1, y1, x2, y2] as 0.0–1.0 fractions of the SLIDE CONTENT AREA only.
  Top-left corner of the slide = [0, 0]. Bottom-right corner = [1, 1].
  Do NOT use full-image coordinates. Always normalize to the slide area.

COLOR RULES — for every text element extract:
- color: closest hex (e.g. "#E07B00" for orange, "#000000" for black)
- background: hex if text sits on a colored box, else null
- bold: true/false
- size: "title" | "subtitle" | "body" | "small" | "tiny"

CHART/IMAGE CROP RULES — when you see a chart, graph, or table:
- type: "image_crop"
- chart_type: "bar" | "line" | "pie" | "table" | "mixed" | "unknown"
- bbox: [x1, y1, x2, y2] fractions of the SLIDE CONTENT AREA (same rule as above)
- chart_title: title text on the chart itself, or null
- description: 1-sentence summary of what the chart shows
- DO NOT extract data values from charts

LAYOUT:
- layout: "single" | "two_column" | "three_column" | "mixed"
- columns: array, index 0 = leftmost column, elements ordered top-to-bottom

INFERENCE:
- If text is partially hidden: add inferred=true, confidence="high"|"medium"|"low"

Return ONLY valid JSON. No markdown fences, no backticks, no explanation text."""

USER_MESSAGE_TEMPLATE = (
    "Analyze this IELTS presentation slide ({width}x{height} pixels).\n"
    "Extract ALL visible content following the system instructions. Return only JSON."
)

RETRY_SUFFIX = "\n\nReturn ONLY the JSON object, nothing else."


def _strip_fences(text: str) -> str:
    """Remove surrounding markdown code fences from a model response."""
    text = text.strip()
    if text.startswith("```"):
        # Drop the opening fence line (``` or ```json) and the closing fence.
        lines = text.splitlines()
        if lines and lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip().startswith("```"):
            lines = lines[:-1]
        text = "\n".join(lines).strip()
    return text


def _call_claude(client, image_b64: str, user_message: str) -> str:
    """Call the Vision API with 429 exponential backoff (max 3 retries)."""
    delays = [2, 4, 8]
    attempt = 0
    while True:
        try:
            resp = client.messages.create(
                model=MODEL,
                max_tokens=MAX_TOKENS,
                system=SYSTEM_PROMPT,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "image",
                                "source": {
                                    "type": "base64",
                                    "media_type": "image/png",
                                    "data": image_b64,
                                },
                            },
                            {"type": "text", "text": user_message},
                        ],
                    }
                ],
            )
            return resp.content[0].text
        except Exception as exc:  # noqa: BLE001
            status = getattr(exc, "status_code", None)
            is_rate_limit = status == 429 or "429" in str(exc)
            if is_rate_limit and attempt < len(delays):
                time.sleep(delays[attempt])
                attempt += 1
                continue
            raise


def analyze_slide(slide: dict, client) -> dict:
    """Extract structured content from one slide via Claude Vision.

    ``slide`` is a SlideMeta dict with a ``representative_frame`` path that
    must be resolvable from the current working directory (callers should
    pass an absolute path, or chdir into the job directory).

    Returns the parsed slide-content dict, with ``slide_id`` ensured present.
    On unrecoverable parse failure returns an error stub.
    """
    sid = slide["slide_id"]
    image_path = Path(slide["representative_frame"])

    with Image.open(image_path) as img:
        width, height = img.size

    with open(image_path, "rb") as fh:
        image_b64 = base64.standard_b64encode(fh.read()).decode("ascii")

    user_message = USER_MESSAGE_TEMPLATE.format(width=width, height=height)

    for attempt in range(2):  # initial try + one retry
        msg = user_message if attempt == 0 else user_message + RETRY_SUFFIX
        try:
            raw = _call_claude(client, image_b64, msg)
            data = json.loads(_strip_fences(raw))
        except json.JSONDecodeError:
            continue  # retry once with the stricter instruction
        data.setdefault("slide_id", sid)
        return data

    print(f"[vision] WARNING: could not parse JSON for {sid}; storing error stub")
    return {"slide_id": sid, "error": True, "columns": []}


def analyze_all(slides: list[dict], client, out_dir: str | Path) -> list[dict]:
    """Run :func:`analyze_slide` for every slide and persist the aggregate.

    Writes ``{out_dir}/slides_extracted.json`` and returns the list.
    """
    out_dir = Path(out_dir)
    results: list[dict] = []
    for slide in slides:
        sid = slide["slide_id"]
        # Resolve the representative frame relative to the job directory.
        resolved = dict(slide)
        resolved["representative_frame"] = str(out_dir / slide["representative_frame"])
        try:
            results.append(analyze_slide(resolved, client))
        except Exception as exc:  # noqa: BLE001
            # Isolate a single slide's failure so the rest of the job survives.
            print(f"[vision] ERROR analyzing {sid}: {type(exc).__name__}: {exc}")
            results.append({"slide_id": sid, "error": True, "columns": []})

    with open(out_dir / "slides_extracted.json", "w", encoding="utf-8") as fh:
        json.dump(results, fh, indent=2)

    return results
