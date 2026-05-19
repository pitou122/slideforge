"""Stage 4 — Image Cropping.

Crop each ``image_crop`` element out of its representative slide frame.
See pipeline_spec.md for the authoritative spec.
"""
from __future__ import annotations

import json
from pathlib import Path

from PIL import Image

MIN_CROP_PX = 50  # skip crops smaller than this in either dimension
PADDING_PX = 4


def crop_images(extracted: list[dict], out_dir: str | Path) -> None:
    """Crop every ``image_crop`` element from its slide's representative PNG.

    ``extracted`` is the list produced by Stage 3. Each cropped region is
    saved under ``{out_dir}/crops/`` and a ``crop_path`` field is added to
    the element dict in place. The updated ``slides_extracted.json`` is
    rewritten so downstream stages see the crop paths.
    """
    out_dir = Path(out_dir)
    crops_dir = out_dir / "crops"
    crops_dir.mkdir(parents=True, exist_ok=True)

    # Map slide_id -> representative frame path (from Stage 2 output).
    slides_detected_path = out_dir / "slides_detected.json"
    rep_by_id: dict[str, str] = {}
    if slides_detected_path.exists():
        with open(slides_detected_path, encoding="utf-8") as fh:
            for s in json.load(fh):
                rep_by_id[s["slide_id"]] = s["representative_frame"]

    for slide in extracted:
        sid = slide.get("slide_id")
        rep = rep_by_id.get(sid)
        if not rep:
            continue
        rep_path = out_dir / rep
        if not rep_path.exists():
            continue

        with Image.open(rep_path) as img:
            W, H = img.size
            slide_img = img.convert("RGB")

            # Convert slide-content-area bbox fractions → full-image pixel coords.
            # Vision returns slide_area=[x1,y1,x2,y2] as fractions of the full
            # image; element bboxes are fractions of that slide content area.
            sa = slide.get("slide_area", [0.0, 0.0, 1.0, 1.0])
            sa_x1, sa_y1, sa_x2, sa_y2 = sa
            sa_px_x = sa_x1 * W
            sa_px_y = sa_y1 * H
            sa_w = (sa_x2 - sa_x1) * W
            sa_h = (sa_y2 - sa_y1) * H

            # Counter for unique crop names — the Vision model does not always
            # return an "id" field on elements, so we cannot rely on it.
            crop_index = 0

            for column in slide.get("columns", []):
                for el in column.get("elements", []):
                    if el.get("type") != "image_crop":
                        continue
                    bbox = el.get("bbox")
                    if not bbox or len(bbox) != 4:
                        continue

                    x1 = max(0, int(sa_px_x + bbox[0] * sa_w) - PADDING_PX)
                    y1 = max(0, int(sa_px_y + bbox[1] * sa_h) - PADDING_PX)
                    x2 = min(W, int(sa_px_x + bbox[2] * sa_w) + PADDING_PX)
                    y2 = min(H, int(sa_px_y + bbox[3] * sa_h) + PADDING_PX)

                    if (x2 - x1) < MIN_CROP_PX or (y2 - y1) < MIN_CROP_PX:
                        continue

                    crop_index += 1
                    el_id = el.get("id") or f"el_{crop_index:03d}"
                    crop_name = f"{sid}_{el_id}.png"
                    crop_rel = f"crops/{crop_name}"
                    slide_img.crop((x1, y1, x2, y2)).save(out_dir / crop_rel)
                    el["crop_path"] = crop_rel

    # Persist the crop_path additions for downstream stages / debugging.
    with open(out_dir / "slides_extracted.json", "w", encoding="utf-8") as fh:
        json.dump(extracted, fh, indent=2)
