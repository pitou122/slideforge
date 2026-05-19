# Pipeline Spec — Video → PPTX

## Stage 1 — Frame Extraction (`src/extractor.py`)

Extract frames from the input video at a configurable sample rate.

**Logic:**
- Open video with `cv2.VideoCapture`
- Calculate frame interval: `interval = int(video_fps / sample_fps)`
- Save every N-th frame as PNG to `{out_dir}/frames/frame_{N:05d}.png`
- Record timestamp for each saved frame

**Output:** `frames_meta.json`
```json
[
  { "frame_id": "frame_00001", "timestamp_sec": 1.0, "path": "frames/frame_00001.png" },
  ...
]
```

---

## Stage 2 — Slide Change Detection (`src/detector.py`)

Compare consecutive frames to find slide transitions.

**Logic:**
- Compare pairs with SSIM (`skimage.metrics.structural_similarity`)
- If SSIM < threshold → new slide detected
- Debounce: skip next 2 seconds of frames after a change (avoids catching mid-transition)
- From each slide's frames, pick the sharpest one:
  `sharpness = cv2.Laplacian(gray, cv2.CV_64F).var()`
- Copy best frame to `{out_dir}/slides/slide_NNN.png`

**Output:** `slides_detected.json`
```json
[
  {
    "slide_id": "slide_001",
    "representative_frame": "slides/slide_001.png",
    "start_sec": 0.0,
    "end_sec": 14.0
  }
]
```

---

## Stage 3 — Claude Vision Extraction (`src/vision.py`)

Send each representative slide image to Claude Vision API and extract structured content.

**API config:**
- Model: `claude-sonnet-4-20250514`
- Max tokens: 4000
- Image: base64-encoded PNG

**System prompt** (use exactly as written):
```
You are an IELTS slide content extractor. Analyze this presentation slide image
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

COLOR RULES — for every text element extract:
- color: closest hex (e.g. "#E07B00" for orange, "#000000" for black)
- background: hex if text sits on a colored box, else null
- bold: true/false
- size: "title" | "subtitle" | "body" | "small" | "tiny"

CHART/IMAGE CROP RULES — when you see a chart, graph, or table:
- type: "image_crop"
- chart_type: "bar" | "line" | "pie" | "table" | "mixed" | "unknown"
- bbox: [x1_pct, y1_pct, x2_pct, y2_pct] as 0.0–1.0 percentages of image size
- chart_title: title text on the chart itself, or null
- description: 1-sentence summary of what the chart shows
- DO NOT extract data values from charts

LAYOUT:
- layout: "single" | "two_column" | "three_column" | "mixed"
- columns: array, index 0 = leftmost column, elements ordered top-to-bottom

INFERENCE:
- If text is partially hidden: add inferred=true, confidence="high"|"medium"|"low"

Return ONLY valid JSON. No markdown fences, no backticks, no explanation text.
```

**User message:**
```
Analyze this IELTS presentation slide ({WIDTH}x{HEIGHT} pixels).
Extract ALL visible content following the system instructions. Return only JSON.
```

**Expected output schema per slide:**
```json
{
  "slide_id": "slide_001",
  "layout": "three_column",
  "background_color": "#ffffff",
  "title": {
    "text": "1/ change over time (standard 4 categories/ ...)",
    "color": "#E07B00",
    "size": "title",
    "bold": true,
    "background": null,
    "inferred": true,
    "confidence": "high"
  },
  "columns": [
    {
      "index": 0,
      "elements": [
        { "id": "el_001", "type": "badge",      "text": "TASK 1", "color": "#ffffff", "background": "#333333", "size": "small", "bold": true },
        { "id": "el_002", "type": "date_label", "text": "13/5/2017", "color": "#666666", "size": "small", "bold": false },
        { "id": "el_003", "type": "paragraph",  "text": "The bar chart below shows...", "color": "#000000", "background": null, "size": "body", "bold": false },
        { "id": "el_004", "type": "image_crop", "chart_type": "bar", "chart_title": "Percentage of spending...", "bbox": [0.01, 0.30, 0.32, 0.72], "description": "Grouped bar chart across 4 countries" },
        { "id": "el_005", "type": "highlight_box", "text": "The graph below shows...", "color": "#0033cc", "background": "#fff3cd", "size": "body", "bold": false }
      ]
    }
  ],
  "excluded": ["webcam_overlay_top_right", "cursor_menu"]
}
```

**Error handling:**
- Strip markdown fences before JSON parsing
- On invalid JSON: retry once with added instruction "Return ONLY the JSON object, nothing else"
- On second failure: log warning, store `{"slide_id": sid, "error": true, "columns": []}`
- On rate limit (429): exponential backoff, max 3 retries (2s, 4s, 8s)

---

## Stage 4 — Image Cropping (`src/cropper.py`)

Crop each `image_crop` element from the representative slide frame.

**Logic:**
- Load representative PNG with PIL
- For each element with `type == "image_crop"`:
  - Convert bbox percentages to pixels with 4px padding:
    ```
    x1 = max(0, int(bbox[0] * W) - 4)
    y1 = max(0, int(bbox[1] * H) - 4)
    x2 = min(W, int(bbox[2] * W) + 4)
    y2 = min(H, int(bbox[3] * H) + 4)
    ```
  - Skip if crop area < 50×50px
  - Save to `{out_dir}/crops/{slide_id}_{el_id}.png`
  - Add `crop_path` field to the element dict

---

## Stage 5 — PPTX Generation (`src/builder.py`)

Rebuild the slide deck using python-pptx.

**Slide dimensions:** 10" × 5.625" (16:9)

**Column layout:**
| layout | col X positions | col widths |
|--------|----------------|------------|
| single | [0.5"] | [9.0"] |
| two_column | [0.5", 5.25"] | [4.5", 4.5"] |
| three_column | [0.5", 3.67", 6.84"] | [2.9", 2.9", 2.9"] |

**Element rendering:**

| type | how to render |
|------|---------------|
| `badge` | Textbox with background fill, bold text |
| `date_label` | Plain textbox, small muted text |
| `paragraph` | Textbox with word wrap, extracted color |
| `highlight_box` | Textbox with background fill color |
| `image_crop` | `add_picture()` scaled to column width, preserve aspect ratio |

**Font sizes:**
| size | pt |
|------|----|
| title | 28 |
| subtitle | 20 |
| body | 12 |
| small | 10 |
| tiny | 8 |

**Title placement:** x=0.3", y=0.15", w=9.4", h=0.9"
**Column content starts at:** y=1.2", stacks downward with 0.1" gap between elements
**Skip element if** its bottom edge would exceed y=5.3" (slide boundary)

---

## Output Directory Structure

```
jobs/{job_id}/
├── input_video.mp4
├── frames/
│   └── frame_00001.png ... frame_NNNNN.png
├── slides/
│   └── slide_001.png ... slide_NNN.png
├── crops/
│   └── slide_001_el_004.png ...
├── frames_meta.json
├── slides_detected.json
├── slides_extracted.json
└── output.pptx              ← final output
```

## CLI Flags (for FastAPI query params)

| param | type | default | description |
|-------|------|---------|-------------|
| fps | float | 1.0 | Frames per second to sample |
| threshold | float | 0.85 | SSIM slide-change threshold |
