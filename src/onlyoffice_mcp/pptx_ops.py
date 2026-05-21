"""PowerPoint operations using python-pptx."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.util import Inches, Pt


# python-pptx ships with the default Office template. Layout indexes:
#   0 = Title Slide        (title + subtitle)
#   1 = Title and Content  (title + body bullets)
#   5 = Title Only         (title at top, rest blank)
#   6 = Blank
_LAYOUT_INDEX = {
    "title": 0,
    "content": 1,
    "title_only": 5,
    "blank": 6,
    "image": 5,
}


def _layout(prs: Presentation, name: str):
    idx = _LAYOUT_INDEX.get(name, 1)
    # Be defensive — some templates have fewer layouts than expected.
    if idx >= len(prs.slide_layouts):
        idx = min(idx, len(prs.slide_layouts) - 1)
    return prs.slide_layouts[idx]


def _set_body_bullets(placeholder, bullets: list[str]) -> None:
    if not bullets:
        return
    tf = placeholder.text_frame
    tf.text = str(bullets[0])
    for line in bullets[1:]:
        p = tf.add_paragraph()
        p.text = str(line)


def create(path: str, slides: list[dict]) -> str:
    """Create a .pptx file at `path`.

    Each slide dict supports:
      - {"layout": "title", "title": str, "subtitle": str}
      - {"layout": "content", "title": str, "body": [bullets]}
      - {"layout": "title_only", "title": str}
      - {"layout": "image", "title": str, "image_path": str,
         "left_inches": float, "top_inches": float, "width_inches": float}
      - {"layout": "blank"}
    """
    out = Path(path).expanduser().resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    prs = Presentation()

    for slide_def in slides:
        layout_name = slide_def.get("layout", "content")
        slide = prs.slides.add_slide(_layout(prs, layout_name))

        # Title is always placeholder index 0 (when present).
        if slide.shapes.title and "title" in slide_def:
            slide.shapes.title.text = str(slide_def["title"])

        if layout_name == "title":
            # Subtitle is placeholder index 1.
            if len(slide.placeholders) > 1 and "subtitle" in slide_def:
                slide.placeholders[1].text = str(slide_def["subtitle"])

        elif layout_name == "content":
            body_items = slide_def.get("body", [])
            if body_items and len(slide.placeholders) > 1:
                _set_body_bullets(slide.placeholders[1], body_items)

        elif layout_name == "image":
            img = slide_def.get("image_path")
            if img:
                img_p = Path(img).expanduser().resolve()
                if not img_p.exists():
                    raise ValueError(f"Image not found: {img_p}")
                left = Inches(slide_def.get("left_inches", 1.0))
                top = Inches(slide_def.get("top_inches", 1.5))
                width = Inches(slide_def.get("width_inches", 8.0))
                slide.shapes.add_picture(str(img_p), left, top, width=width)

        # Add a notes block if requested.
        if slide_def.get("notes"):
            notes_tf = slide.notes_slide.notes_text_frame
            notes_tf.text = str(slide_def["notes"])

    prs.save(str(out))
    return str(out)


def read(path: str) -> dict:
    """Extract slide titles and body text from a .pptx file."""
    in_ = Path(path).expanduser().resolve()
    if not in_.exists():
        raise FileNotFoundError(in_)
    prs = Presentation(str(in_))
    result: dict[str, Any] = {"slide_count": len(prs.slides), "slides": []}
    for i, slide in enumerate(prs.slides, 1):
        info: dict[str, Any] = {"index": i, "title": "", "text": [], "notes": ""}
        title_shape = slide.shapes.title
        if title_shape and title_shape.has_text_frame:
            info["title"] = title_shape.text_frame.text
        for shape in slide.shapes:
            if shape is title_shape:
                continue
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                if paragraph.text:
                    info["text"].append(paragraph.text)
        if slide.has_notes_slide:
            info["notes"] = slide.notes_slide.notes_text_frame.text
        result["slides"].append(info)
    return result


def add_slide(path: str, slide_def: dict) -> str:
    """Append a single slide to an existing .pptx file."""
    in_ = Path(path).expanduser().resolve()
    if not in_.exists():
        raise FileNotFoundError(in_)
    prs = Presentation(str(in_))

    layout_name = slide_def.get("layout", "content")
    slide = prs.slides.add_slide(_layout(prs, layout_name))

    if slide.shapes.title and "title" in slide_def:
        slide.shapes.title.text = str(slide_def["title"])
    if layout_name == "title" and len(slide.placeholders) > 1 and "subtitle" in slide_def:
        slide.placeholders[1].text = str(slide_def["subtitle"])
    elif layout_name == "content":
        body_items = slide_def.get("body", [])
        if body_items and len(slide.placeholders) > 1:
            _set_body_bullets(slide.placeholders[1], body_items)
    elif layout_name == "image" and slide_def.get("image_path"):
        img_p = Path(slide_def["image_path"]).expanduser().resolve()
        if not img_p.exists():
            raise ValueError(f"Image not found: {img_p}")
        slide.shapes.add_picture(
            str(img_p),
            Inches(slide_def.get("left_inches", 1.0)),
            Inches(slide_def.get("top_inches", 1.5)),
            width=Inches(slide_def.get("width_inches", 8.0)),
        )

    prs.save(str(in_))
    return str(in_)
