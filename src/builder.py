"""Stage 5 — PPTX Generation.

Rebuild the slide deck from the extracted JSON using python-pptx.
Element positions come directly from the Vision-returned bbox (fractions of
the slide content area), mapped onto the 10"×5.625" PPTX canvas.
"""
from __future__ import annotations

import math
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Inches, Pt

# Slide is 10" x 5.625" (16:9).
SLIDE_W = 10.0
SLIDE_H = 5.625

FONT_PT = {"title": 28, "subtitle": 20, "body": 12, "small": 10, "tiny": 8}

# Fallback geometry used only when an element has no bbox.
TITLE_BOX = {"x": 0.3, "y": 0.15, "w": 9.4, "h": 0.9}
CONTENT_TOP = 1.2
ELEMENT_GAP = 0.1
BOTTOM_LIMIT = 5.3

LAYOUTS = {
    "single":       {"x": [0.5],        "w": [9.0]},
    "two_column":   {"x": [0.5, 5.25],  "w": [4.5, 4.5]},
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


def _resolve_layout(layout: str, columns: list) -> dict:
    if columns:
        n = max(
            len(columns),
            max((c.get("index", 0) for c in columns), default=0) + 1,
        )
    else:
        n = {"two_column": 2, "three_column": 3}.get(layout, 1)
    if n >= 3:
        return LAYOUTS["three_column"]
    if n == 2:
        return LAYOUTS["two_column"]
    return LAYOUTS["single"]


def _bbox_to_rect(bbox: list) -> tuple[float, float, float, float]:
    """Convert bbox fractions of slide area → inches on the PPTX canvas."""
    x1, y1, x2, y2 = bbox
    x = x1 * SLIDE_W
    y = y1 * SLIDE_H
    w = max(0.1, (x2 - x1) * SLIDE_W)
    h = max(0.05, (y2 - y1) * SLIDE_H)
    return x, y, w, h


def _estimate_text_height(text: str, font_pt: int, width_in: float) -> float:
    text = text or ""
    char_w_in = (font_pt * 0.55) / 72.0
    chars_per_line = max(1, int(width_in / char_w_in))
    explicit_lines = text.count("\n") + 1
    wrapped_lines = sum(
        max(1, math.ceil(len(seg) / chars_per_line)) for seg in text.split("\n")
    )
    lines = max(explicit_lines, wrapped_lines)
    line_h_in = (font_pt * 1.25) / 72.0
    return lines * line_h_in + 0.12


def _add_text_at(slide, el: dict, x: float, y: float, w: float, h: float,
                 *, default_color: str) -> None:
    text = el.get("text", "") or ""
    size_key = el.get("size", "body")
    font_pt = FONT_PT.get(size_key, FONT_PT["body"])

    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP

    para = tf.paragraphs[0]
    run = para.add_run()
    run.text = text
    run.font.size = Pt(font_pt)
    run.font.bold = bool(el.get("bold", False))
    run.font.color.rgb = _rgb(el.get("color"), default_color)

    bg = el.get("background")
    if bg:
        box.fill.solid()
        box.fill.fore_color.rgb = _rgb(bg, "FFFFFF")


def _add_image_at(slide, el: dict, out_dir: Path,
                  x: float, y: float, w: float, h: float) -> None:
    crop_path = el.get("crop_path")
    if not crop_path:
        return
    full = out_dir / crop_path
    if not full.exists():
        return
    from PIL import Image as _Image
    with _Image.open(full) as im:
        iw, ih = im.size
    if not iw or not ih:
        return
    # Fit within the bbox while preserving aspect ratio.
    aspect = iw / ih
    fit_w = min(w, h * aspect)
    fit_h = fit_w / aspect
    slide.shapes.add_picture(str(full), Inches(x), Inches(y),
                             width=Inches(fit_w), height=Inches(fit_h))


def build_pptx(extracted: list[dict], out_dir: str | Path) -> Path:
    out_dir = Path(out_dir)

    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    blank_layout = prs.slide_layouts[6]

    for slide_data in extracted:
        slide = prs.slides.add_slide(blank_layout)

        bg_hex = slide_data.get("background_color")
        if bg_hex:
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = _rgb(bg_hex, "FFFFFF")

        # --- Title ---
        title = slide_data.get("title")
        if title and title.get("text"):
            bbox = title.get("bbox")
            if bbox and len(bbox) == 4:
                tx, ty, tw, th = _bbox_to_rect(bbox)
            else:
                tx, ty, tw, th = (TITLE_BOX["x"], TITLE_BOX["y"],
                                  TITLE_BOX["w"], TITLE_BOX["h"])
            box = slide.shapes.add_textbox(
                Inches(tx), Inches(ty), Inches(tw), Inches(th))
            tf = box.text_frame
            tf.word_wrap = True
            run = tf.paragraphs[0].add_run()
            run.text = title["text"]
            run.font.size = Pt(FONT_PT.get(title.get("size", "title"), 28))
            run.font.bold = bool(title.get("bold", True))
            run.font.color.rgb = _rgb(title.get("color"), "E07B00")

        # --- Column elements ---
        columns = slide_data.get("columns", []) or []
        geom = _resolve_layout(slide_data.get("layout", "single"), columns)
        n_cols = len(geom["x"])

        for column in columns:
            idx = min(column.get("index", 0), n_cols - 1)
            col_x = geom["x"][idx]
            col_w = geom["w"][idx]
            cursor_y = CONTENT_TOP  # fallback cursor for bbox-less elements

            for el in column.get("elements", []):
                el_type = el.get("type")
                bbox = el.get("bbox")

                if bbox and len(bbox) == 4:
                    # Primary path: place element exactly at its bbox position.
                    ex, ey, ew, eh = _bbox_to_rect(bbox)
                    if el_type == "image_crop":
                        _add_image_at(slide, el, out_dir, ex, ey, ew, eh)
                    else:
                        _add_text_at(slide, el, ex, ey, ew, eh,
                                     default_color="000000")
                    cursor_y = ey + eh + ELEMENT_GAP
                else:
                    # Fallback: stack with cursor (no bbox returned by Vision).
                    if el_type == "image_crop":
                        crop_path = el.get("crop_path")
                        if not crop_path or not (out_dir / crop_path).exists():
                            continue
                        from PIL import Image as _Image
                        with _Image.open(out_dir / crop_path) as im:
                            iw, ih = im.size
                        if not iw:
                            continue
                        max_h = BOTTOM_LIMIT - cursor_y - 0.05
                        if max_h < 0.2:
                            continue
                        h_at_full_w = col_w * ih / iw
                        if h_at_full_w <= max_h:
                            slide.shapes.add_picture(
                                str(out_dir / crop_path),
                                Inches(col_x), Inches(cursor_y),
                                width=Inches(col_w))
                            height = h_at_full_w
                        else:
                            slide.shapes.add_picture(
                                str(out_dir / crop_path),
                                Inches(col_x), Inches(cursor_y),
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
                        _add_text_at(slide, el, col_x, cursor_y, col_w, height,
                                     default_color="000000")
                        cursor_y += height + ELEMENT_GAP

    out_path = out_dir / "output.pptx"
    prs.save(str(out_path))
    return out_path
