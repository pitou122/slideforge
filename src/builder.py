"""Stage 5 — PPTX Generation.

Column layout is inferred from bbox.x1 positions so left/right split is
correct even when Vision reports layout="single". Within each column,
elements are stacked top-to-bottom with cursor_y so images fill the column
width instead of being crammed into exact-bbox corner positions.
"""
from __future__ import annotations

import math
from pathlib import Path

from PIL import Image as _PIL_Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Inches, Pt

SLIDE_W = 10.0
SLIDE_H = 5.625

FONT_PT = {"title": 28, "subtitle": 20, "body": 12, "small": 10, "tiny": 8}

TITLE_BOX = {"x": 0.3, "y": 0.15, "w": 9.4, "h": 0.9}
CONTENT_TOP = 1.2
ELEMENT_GAP = 0.1
BOTTOM_LIMIT = 5.3
# Image bbox area < 2% of slide area → logo/watermark, skip.
MIN_BBOX_AREA = 0.02

LAYOUTS = {
    "single":       {"x": [0.5],            "w": [9.0]},
    "two_column":   {"x": [0.5, 5.25],      "w": [4.5, 4.5]},
    "three_column": {"x": [0.5, 3.67, 6.84], "w": [2.9, 2.9, 2.9]},
}


def _rgb(hex_str: str | None, default: str = "000000") -> RGBColor:
    s = (hex_str or "").lstrip("#").strip()
    if len(s) == 6:
        try:
            return RGBColor.from_string(s.upper())
        except ValueError:
            pass
    return RGBColor.from_string(default)


def _estimate_text_height(text: str, font_pt: int, width_in: float) -> float:
    text = text or ""
    char_w_in = (font_pt * 0.55) / 72.0
    chars_per_line = max(1, int(width_in / char_w_in))
    explicit_lines = text.count("\n") + 1
    wrapped_lines = sum(
        max(1, math.ceil(len(seg) / chars_per_line)) for seg in text.split("\n")
    )
    lines = max(explicit_lines, wrapped_lines)
    return lines * (font_pt * 1.25) / 72.0 + 0.12


def _add_text_block(slide, el: dict, x: float, y: float,
                    w: float, h: float, *, default_color: str) -> None:
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    run = tf.paragraphs[0].add_run()
    run.text = el.get("text", "") or ""
    font_pt = FONT_PT.get(el.get("size", "body"), FONT_PT["body"])
    run.font.size = Pt(font_pt)
    run.font.bold = bool(el.get("bold", False))
    run.font.color.rgb = _rgb(el.get("color"), default_color)
    bg = el.get("background")
    if bg:
        box.fill.solid()
        box.fill.fore_color.rgb = _rgb(bg, "FFFFFF")


def _infer_geom(all_elements: list[tuple[int, dict]]) -> dict:
    """Pick column geometry from bbox.x1 positions.

    Only infers two_column (never three) — small badges/dates near x1=0.8+
    are part of the right column, not a third column.  Three-column layouts
    must be declared by Vision via the layout field and handled upstream.
    """
    x1_vals = [
        el["bbox"][0]
        for _, el in all_elements
        if el.get("bbox") and len(el["bbox"]) == 4
    ]
    if not x1_vals:
        return LAYOUTS["single"]
    # A clear left/right split: at least one element clearly on each side.
    has_left  = any(x < 0.45 for x in x1_vals)
    has_right = any(x > 0.55 for x in x1_vals)
    if has_left and has_right:
        return LAYOUTS["two_column"]
    return LAYOUTS["single"]


def build_pptx(extracted: list[dict], out_dir: str | Path) -> Path:
    out_dir = Path(out_dir)
    prs = Presentation()
    prs.slide_width  = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    blank_layout = prs.slide_layouts[6]

    for slide_data in extracted:
        slide = prs.slides.add_slide(blank_layout)

        bg_hex = slide_data.get("background_color")
        if bg_hex:
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = _rgb(bg_hex, "FFFFFF")

        # --- Top-level title (some Vision responses include this) ---
        title = slide_data.get("title")
        if title and title.get("text"):
            box = slide.shapes.add_textbox(
                Inches(TITLE_BOX["x"]), Inches(TITLE_BOX["y"]),
                Inches(TITLE_BOX["w"]), Inches(TITLE_BOX["h"]))
            tf = box.text_frame
            tf.word_wrap = True
            run = tf.paragraphs[0].add_run()
            run.text = title["text"]
            run.font.size = Pt(FONT_PT.get(title.get("size", "title"), 28))
            run.font.bold = bool(title.get("bold", True))
            run.font.color.rgb = _rgb(title.get("color"), "E07B00")

        # --- Flatten all column elements ---
        columns = slide_data.get("columns", []) or []
        all_els: list[tuple[int, dict]] = [
            (col.get("index", 0), el)
            for col in columns
            for el in col.get("elements", [])
        ]
        if not all_els:
            continue

        # --- Infer column geometry from bbox x1 ---
        geom = _infer_geom(all_els)
        n_cols = len(geom["x"])

        # --- Assign each element to a column by bbox.x1 ---
        # Elements without bbox fall back to their original column index.
        col_buckets: list[list[tuple[float, dict]]] = [[] for _ in range(n_cols)]
        for orig_idx, el in all_els:
            bbox = el.get("bbox")
            if bbox and len(bbox) == 4:
                col = min(int(bbox[0] * n_cols), n_cols - 1)
                sort_y = bbox[1]
            else:
                col = min(orig_idx, n_cols - 1)
                sort_y = 0.0
            col_buckets[col].append((sort_y, el))

        # Sort each column top-to-bottom.
        for bucket in col_buckets:
            bucket.sort(key=lambda t: t[0])

        # --- Render columns with cursor stacking ---
        for c_idx, items in enumerate(col_buckets):
            col_x = geom["x"][c_idx]
            col_w = geom["w"][c_idx]
            cursor_y = CONTENT_TOP

            for _, el in items:
                el_type = el.get("type")

                if el_type == "image_crop":
                    # Skip logos/watermarks — bbox area < 2% of slide area.
                    el_bbox = el.get("bbox") or []
                    if len(el_bbox) == 4:
                        bbox_area = (el_bbox[2] - el_bbox[0]) * (el_bbox[3] - el_bbox[1])
                        if bbox_area < MIN_BBOX_AREA:
                            continue
                    crop_path = el.get("crop_path")
                    if not crop_path:
                        continue
                    full = out_dir / crop_path
                    if not full.exists():
                        continue
                    with _PIL_Image.open(full) as im:
                        iw, ih = im.size
                    if not iw or not ih:
                        continue
                    max_h = BOTTOM_LIMIT - cursor_y - 0.05
                    if max_h < 0.3:
                        continue
                    h_at_col_w = col_w * ih / iw
                    if h_at_col_w <= max_h:
                        slide.shapes.add_picture(
                            str(full), Inches(col_x), Inches(cursor_y),
                            width=Inches(col_w))
                        height = h_at_col_w
                    else:
                        slide.shapes.add_picture(
                            str(full), Inches(col_x), Inches(cursor_y),
                            height=Inches(max_h))
                        height = max_h
                    cursor_y += height + ELEMENT_GAP

                else:
                    size_key = el.get("size", "body")
                    font_pt = FONT_PT.get(size_key, FONT_PT["body"])
                    height = _estimate_text_height(
                        el.get("text", ""), font_pt, col_w)
                    if cursor_y + height > BOTTOM_LIMIT:
                        continue
                    _add_text_block(slide, el, col_x, cursor_y,
                                    col_w, height, default_color="000000")
                    cursor_y += height + ELEMENT_GAP

    out_path = out_dir / "output.pptx"
    prs.save(str(out_path))
    return out_path
