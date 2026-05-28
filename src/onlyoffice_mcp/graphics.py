"""Standalone graphic / infographic generation for slide-style reports.

These tools produce transparent (or solid) PNG images that can be composited
into documents via ``docx_set_background_image`` or embedded as ``image``
blocks. Unlike :mod:`charts` (which embeds charts *inside* a document on a
white canvas), everything here renders on a transparent background with
light text so it sits on a dark themed page.

Quality: all PIL renderers draw on a 2x supersampled canvas and downscale with
LANCZOS for smooth anti-aliased edges. Cards use vertical gradient fills with a
top highlight and an optional left accent stripe; circular badges/nodes get a
soft top "gloss"; the background carries a corner vignette. matplotlib charts
stroke their labels for legibility on any colour.

Capabilities:
- ``tech_background``    dark gradient + hexagon grid + glow + vignette + optional header band
- ``recolor_image``      recolour a logo/mark to a flat colour on transparency + tight crop
- ``key_logo``           key a logo's flat background to transparency, KEEPING its colours
- ``donut_chart``        ring chart with centre label (transparent, stroked light text)
- ``bar_chart``          horizontal/vertical bars with value labels (transparent, stroked text)
- ``bubble_cards``       rounded "bubble" card grid (ID badge + title + impact + severity pill)
- ``node_infographic``   glossy circular value nodes on a dashed zigzag connector
- ``numbered_cards``     numbered 2-column recommendation/step cards
- ``decorative_panel``   abstract hexagon cluster with an optional lock/shield motif

All renderers are deterministic (fixed RNG seed) so re-runs are reproducible.
"""

from __future__ import annotations

import logging
import math
import random
from pathlib import Path

import numpy as np
import matplotlib

matplotlib.use("Agg")  # noqa: E402  (headless backend, before pyplot)
import matplotlib.pyplot as plt  # noqa: E402
import matplotlib.patheffects as pe  # noqa: E402
from PIL import Image, ImageDraw, ImageFilter, ImageFont  # noqa: E402

from .validation import (
    validate_path, validate_color, validate_choice, validate_records, validate_bounded_int,
)
from . import safety

log = logging.getLogger(__name__)

_FONT_DIR = matplotlib.get_data_path() + "/fonts/ttf"
_SS = 2  # supersampling factor for anti-aliasing
_CARD_TOP = (20, 50, 88)
_CARD_BOT = (11, 30, 56)
_CARD_OUTLINE = (54, 120, 192, 175)


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------

def _ss_for(w: int, h: int) -> int:
    """Supersample factor, dropped to 1 if the scaled canvas would be huge."""
    return _SS if (w * h * _SS * _SS) <= 34_000_000 else 1


def _font(size: float, bold: bool = True) -> ImageFont.FreeTypeFont:
    name = "DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf"
    return ImageFont.truetype(f"{_FONT_DIR}/{name}", max(1, int(size)))


def _rgb(color: str) -> tuple[int, int, int]:
    h = validate_color(color)
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _wrap(draw, text, font, max_w):
    words = str(text).split()
    lines, cur = [], ""
    for w in words:
        t = (cur + " " + w).strip()
        if draw.textlength(t, font=font) <= max_w:
            cur = t
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines or [""]


def _ctext(draw, cx, cy, text, font, fill):
    b = draw.textbbox((0, 0), str(text), font=font)
    draw.text((cx - (b[2] - b[0]) / 2 - b[0], cy - (b[3] - b[1]) / 2 - b[1]), str(text), font=font, fill=fill)


def _hexagon(cx, cy, r):
    return [(cx + r * math.cos(math.pi / 180 * (60 * i - 30)),
             cy + r * math.sin(math.pi / 180 * (60 * i - 30))) for i in range(6)]


def _vgrad(w, h, top, bot, alpha=255):
    """Fast vertical-gradient RGBA image via numpy."""
    w, h = max(1, int(w)), max(1, int(h))
    t = np.linspace(0.0, 1.0, h)[:, None]
    arr = np.empty((h, w, 4), np.uint8)
    for i in range(3):
        arr[:, :, i] = np.clip(top[i] + (bot[i] - top[i]) * t, 0, 255).astype(np.uint8)
    arr[:, :, 3] = alpha
    return Image.fromarray(arr, "RGBA")


def _round_mask(w, h, radius, left_cut=0):
    m = Image.new("L", (w, h), 0)
    ImageDraw.Draw(m).rounded_rectangle([0, 0, w - 1, h - 1], radius=radius, fill=255)
    if left_cut > 0:
        ImageDraw.Draw(m).rectangle([0, 0, left_cut, h], fill=0)
    return m


def _card(base, box, radius, *, accent=None, stripe=0, outline=_CARD_OUTLINE, ow=2,
          top=_CARD_TOP, bot=_CARD_BOT, alpha=250):
    """Draw a rounded card with a vertical gradient fill, top highlight and an
    optional left accent stripe (the accent colour shows through a left band)."""
    x0, y0, x1, y1 = (int(v) for v in box)
    w, h = x1 - x0, y1 - y0
    d = ImageDraw.Draw(base)
    if accent and stripe > 0:
        d.rounded_rectangle([x0, y0, x1, y1], radius=radius, fill=(*accent, 255))
    grad = _vgrad(w, h, top, bot, alpha)
    base.paste(grad, (x0, y0), _round_mask(w, h, radius, left_cut=stripe if (accent and stripe > 0) else 0))
    if outline:
        d.rounded_rectangle([x0, y0, x1, y1], radius=radius, outline=outline, width=ow)
    # subtle top highlight
    d.line([(x0 + radius, y0 + ow), (x1 - radius, y0 + ow)], fill=(255, 255, 255, 30), width=max(1, ow))


def _gloss(base, cx, cy, r):
    """Soft top highlight on a circle for a glossy look."""
    g = Image.new("RGBA", base.size, (0, 0, 0, 0))
    ImageDraw.Draw(g).ellipse([cx - r * 0.66, cy - r * 0.92, cx + r * 0.66, cy - r * 0.04], fill=(255, 255, 255, 46))
    base.alpha_composite(g.filter(ImageFilter.GaussianBlur(max(1, r * 0.12))))


def _out(path):
    return validate_path(path, expected_ext="png", for_creation=True, operation="create_graphic")


def _check_dims(width_px, height_px):
    validate_bounded_int(int(width_px), "width_px", min_val=64, max_val=8000)
    validate_bounded_int(int(height_px), "height_px", min_val=64, max_val=8000)


def _src_image(image_path):
    img = Path(image_path).expanduser().resolve()
    safety.check_path_safety(img)
    if not img.exists():
        raise ValueError(f"Image not found: {img}\nProvide a valid path to a PNG/JPG image.")
    ext = img.suffix.lstrip(".").lower()
    if ext not in safety.ALLOWED_IMAGE_EXTENSIONS:
        raise ValueError(f"Unsupported image format '.{ext}'. Allowed: {', '.join(sorted(safety.ALLOWED_IMAGE_EXTENSIONS))}")
    safety.check_file_size(img)
    return img


def _hex_grid(img, accent, r):
    W, H = img.size
    layer = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    hd = ImageDraw.Draw(layer)
    dx, dy = math.sqrt(3) * r, 1.5 * r
    rng = random.Random(7)
    row, y = 0, -r
    while y < H + r:
        x = -r + ((dx / 2) if row % 2 else 0)
        while x < W + r:
            line = (*accent, 70) if rng.random() < 0.06 else (90, 150, 220, 26)
            hd.polygon(_hexagon(x, y, r), outline=line)
            x += dx
        y += dy
        row += 1
    for _ in range(18):
        hd.polygon(_hexagon(rng.randint(0, W), rng.randint(0, H), r), fill=(*accent, rng.randint(8, 20)))
    return Image.alpha_composite(img, layer)


def _vignette(img, strength=0.55):
    """Darken the corners with a radial mask for depth."""
    W, H = img.size
    yy, xx = np.mgrid[0:H, 0:W]
    cx, cy = W / 2, H / 2
    d = np.sqrt(((xx - cx) / cx) ** 2 + ((yy - cy) / cy) ** 2)
    v = np.clip((d - 0.55) / 0.85, 0, 1) * strength
    overlay = np.zeros((H, W, 4), np.uint8)
    overlay[:, :, 3] = (v * 255).astype(np.uint8)
    return Image.alpha_composite(img, Image.fromarray(overlay, "RGBA"))


# --------------------------------------------------------------------------
# 1. Tech background
# --------------------------------------------------------------------------

def tech_background(out_path, *, width_px=2560, height_px=1440, top_color="#06122a",
                    bottom_color="#0a264a", accent_color="#5aa0e0", hexagons=True,
                    glow=True, dot_wave=True, header_text=None, header_subtext=None,
                    header_monogram=None, logo_path=None, logo_on_right=True):
    """Render a dark gradient "tech" background (default 16:9) with an optional
    hexagon grid, corner glow, dotted wave, corner vignette and an optional
    branded header band. Use as a full-bleed document background.
    Returns the output path."""
    out = _out(out_path)
    top, bot, accent = _rgb(top_color), _rgb(bottom_color), _rgb(accent_color)
    W0, H0 = int(width_px), int(height_px)
    if not (16 <= W0 <= 8000 and 16 <= H0 <= 8000):
        raise ValueError("width_px/height_px must be between 16 and 8000.")
    S = _ss_for(W0, H0)
    W, H = W0 * S, H0 * S

    bg = _vgrad(W, H, top, bot, 255).convert("RGBA")

    if glow:
        g = Image.new("L", (W, H), 0)
        ImageDraw.Draw(g).ellipse([W * 0.45, -H * 0.35, W * 1.25, H * 0.65], fill=72)
        g = g.filter(ImageFilter.GaussianBlur(int(H * 0.16)))
        bg = Image.alpha_composite(bg, Image.merge("RGBA", (
            Image.new("L", (W, H), 40), Image.new("L", (W, H), 120), Image.new("L", (W, H), 200), g)))

    if hexagons:
        bg = _hex_grid(bg, accent, r=98 * S)

    if dot_wave:
        dots = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        dd = ImageDraw.Draw(dots)
        rng = random.Random(7)
        for _ in range(int(W0 * H0 / 3850)):
            x = rng.randint(0, W)
            yy = int(H * 0.80 + math.sin(x / (135.0 * S)) * (H * 0.05) + rng.randint(-40 * S, 40 * S))
            if 0 <= yy < H:
                s = rng.choice([2, 2, 3]) * S
                dd.ellipse([x, yy, x + s, yy + s], fill=(120, 200, 255, rng.randint(20, 90)))
        bg = Image.alpha_composite(bg, dots)

    bg = _vignette(bg, 0.5)

    if header_text or logo_path:
        band_h = int(H * 0.072)
        band = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        bdd = ImageDraw.Draw(band)
        bdd.rectangle([0, 0, W, band_h], fill=(4, 12, 28, 150))
        bdd.line([(0, band_h), (W, band_h)], fill=(*accent, 170), width=2 * S)
        bg = Image.alpha_composite(bg, band)
        d = ImageDraw.Draw(bg)
        tx = 70 * S
        if header_monogram:
            chip = int(band_h * 0.44)
            cy = (band_h - chip) // 2
            d.rounded_rectangle([70 * S, cy, 70 * S + chip, cy + chip], radius=8 * S, outline=(120, 200, 255), width=3 * S)
            _ctext(d, 70 * S + chip / 2, cy + chip / 2, header_monogram, _font(chip * 0.42), (255, 255, 255))
            tx = 70 * S + chip + 18 * S
        if header_text:
            d.text((tx, band_h * 0.22), header_text, font=_font(band_h * 0.22), fill=(255, 255, 255))
        if header_subtext:
            d.text((tx, band_h * 0.55), header_subtext, font=_font(band_h * 0.17, bold=False), fill=(159, 196, 232))
        if logo_path:
            logo = Image.open(_src_image(logo_path)).convert("RGBA")
            lw = int(W * 0.16)
            lh = int(logo.size[1] * lw / logo.size[0])
            bg.alpha_composite(logo.resize((lw, lh), Image.LANCZOS),
                               ((W - lw - 70 * S) if logo_on_right else 70 * S, (band_h - lh) // 2))

    if S > 1:
        bg = bg.resize((W0, H0), Image.LANCZOS)
    bg.convert("RGB").save(str(out))
    log.info("tech_background -> %s (%dx%d, ss=%d)", out.name, W0, H0, S)
    return str(out)


# --------------------------------------------------------------------------
# 2. Recolour image
# --------------------------------------------------------------------------

def recolor_image(image_path, out_path, *, color="#FFFFFF", bright_threshold=232,
                  crop=True, pad=14, scale=3):
    """Recolour a logo/mark to a single flat ``color`` on a transparent canvas:
    the white backing becomes transparent and the ink is repainted (alpha by
    darkness, smooth edges), then optionally tight-cropped and upscaled.
    Returns the output path."""
    out = _out(out_path)
    src = Image.open(_src_image(image_path)).convert("RGBA")
    rc = _rgb(color)
    arr = np.asarray(src).astype(np.int16)
    bright = arr[:, :, :3].mean(axis=2)
    a = arr[:, :, 3]
    thr = max(1, min(255, int(bright_threshold)))
    ink = (a > 0) & (bright <= thr)
    alpha = np.clip((thr - bright) / thr * 255 * 1.6, 0, 255).astype(np.uint8)
    alpha = np.where(ink, np.maximum(alpha, 40), 0).astype(np.uint8)
    out_arr = np.zeros_like(arr, dtype=np.uint8)
    out_arr[:, :, 0], out_arr[:, :, 1], out_arr[:, :, 2] = rc
    out_arr[:, :, 3] = alpha
    dst = Image.fromarray(out_arr, "RGBA")
    if crop:
        bbox = dst.getbbox()
        if bbox:
            c = dst.crop(bbox)
            dst = Image.new("RGBA", (c.size[0] + pad * 2, c.size[1] + pad * 2), (rc[0], rc[1], rc[2], 0))
            dst.alpha_composite(c, (pad, pad))
    scale = max(1, min(6, int(scale)))
    if scale > 1:
        dst = dst.resize((dst.size[0] * scale, dst.size[1] * scale), Image.LANCZOS)
    dst.save(str(out))
    log.info("recolor_image -> %s (%dx%d)", out.name, *dst.size)
    return str(out)


def key_logo(image_path, out_path, *, thresh=70, crop=True, pad=8, scale=1, feather=0.0):
    """Key a logo's flat background to transparency while PRESERVING the logo's
    original full colour and interior detail.

    Unlike :func:`recolor_image` (which flattens the mark to one flat colour),
    this keeps the artwork as-is and only removes the surrounding backing — ideal
    for dropping a real colour logo/crest onto any (e.g. dark) background. The
    background is found by flood-filling inward from the four corners and edge
    midpoints, so light areas *enclosed* by darker ink (a white roundel inside a
    crest, white highlights inside a gold crown) are kept rather than punched out.

    ``thresh`` is the flood colour tolerance (sum of per-channel abs diff from the
    sampled edge colour; higher removes more, default 70). ``feather`` softly
    blurs the cut edge by N px (0 = crisp). ``crop`` tight-crops to the kept
    content with ``pad`` transparent px; ``scale`` upsamples 1–6x. Returns path."""
    out = _out(out_path)
    src = Image.open(_src_image(image_path)).convert("RGB")
    W, H = src.size

    # choose a sentinel fill colour that does not already occur in the image
    used = set(src.getdata())
    key = next(((r, g, b) for (r, g, b) in
                [(255, 0, 255), (0, 255, 0), (255, 0, 0), (0, 0, 255), (255, 255, 0), (1, 254, 2)]
                if (r, g, b) not in used), (255, 0, 255))

    flood = src.copy()
    t = max(0, min(765, int(thresh)))
    seeds = [(0, 0), (W - 1, 0), (0, H - 1), (W - 1, H - 1),
             (W // 2, 0), (W // 2, H - 1), (0, H // 2), (W - 1, H // 2)]
    for s in seeds:
        ImageDraw.floodfill(flood, s, key, thresh=t)

    out_arr = np.asarray(src.convert("RGBA")).copy()
    bg_mask = np.all(np.asarray(flood) == np.array(key, np.uint8), axis=2)
    out_arr[bg_mask, 3] = 0
    dst = Image.fromarray(out_arr, "RGBA")

    if feather and feather > 0:
        alpha = dst.split()[3].filter(ImageFilter.GaussianBlur(float(feather)))
        dst.putalpha(alpha)

    if crop:
        bbox = dst.getbbox()
        if bbox:
            c = dst.crop(bbox)
            dst = Image.new("RGBA", (c.size[0] + pad * 2, c.size[1] + pad * 2), (0, 0, 0, 0))
            dst.alpha_composite(c, (pad, pad))

    scale = max(1, min(6, int(scale)))
    if scale > 1:
        dst = dst.resize((dst.size[0] * scale, dst.size[1] * scale), Image.LANCZOS)
    dst.save(str(out))
    log.info("key_logo -> %s (%dx%d, thresh=%d)", out.name, dst.size[0], dst.size[1], t)
    return str(out)


# --------------------------------------------------------------------------
# 3. Charts on transparency
# --------------------------------------------------------------------------

_STROKE = [pe.withStroke(linewidth=3, foreground="#07142b")]


def _fmt_value(v):
    if v >= 1_000_000:
        return f"{v/1_000_000:g}M"
    if v >= 1_000:
        return f"{v/1_000:g}K"
    return f"{v:g}"


def donut_chart(out_path, segments, *, center_text=None, center_subtext=None,
                hole=0.42, show_values=True, dpi=200):
    """Render a ring/donut chart on transparency with stroked light labels.
    ``segments`` = ``[{"label","value","color"}, ...]``; ``center_text`` /
    ``center_subtext`` are drawn in the hole. Returns the output path."""
    out = _out(out_path)
    validate_records(segments, "segments", required=("value",), numeric=("value",),
                     example='{"label": "Critical", "value": 26, "color": "#C0202A"}')
    if not 0.05 <= float(hole) <= 0.95:
        raise ValueError(f"hole must be between 0.05 and 0.95 (ring thickness), got {hole!r}.")
    validate_bounded_int(int(dpi), "dpi", min_val=36, max_val=600)
    vals = [float(s["value"]) for s in segments]
    cols = ["#" + validate_color(s.get("color", "#27B7E6")) for s in segments]
    fig, ax = plt.subplots(figsize=(6.0, 5.0), dpi=int(dpi))
    fig.patch.set_alpha(0)
    ax.set_facecolor("none")
    wedges, _ = ax.pie(vals, colors=cols, startangle=90, counterclock=False,
                       wedgeprops=dict(width=max(0.1, min(0.9, hole)), edgecolor="#07142b", linewidth=4))
    if center_text:
        ax.text(0, 0.12, str(center_text), ha="center", va="center", color="white",
                fontsize=54, fontweight="bold", path_effects=_STROKE)
    if center_subtext:
        ax.text(0, -0.22, str(center_subtext), ha="center", va="center", color="#9fc4e8",
                fontsize=14, fontweight="bold")
    if show_values:
        for w, v in zip(wedges, vals):
            ang = math.radians((w.theta1 + w.theta2) / 2)
            ax.text(0.79 * math.cos(ang), 0.79 * math.sin(ang), _fmt_value(v), ha="center", va="center",
                    color="white", fontsize=20, fontweight="bold", path_effects=_STROKE)
    ax.set(aspect="equal")
    ax.axis("off")
    plt.tight_layout(pad=0)
    fig.savefig(str(out), transparent=True, bbox_inches="tight")
    plt.close(fig)
    log.info("donut_chart -> %s (%d segments)", out.name, len(segments))
    return str(out)


def bar_chart(out_path, bars, *, horizontal=True, log_scale=False, axis_label=None, dpi=200):
    """Render a bar chart on transparency with stroked light labels and value
    annotations. ``bars`` = ``[{"label","value","color"}, ...]``. Set
    ``log_scale`` true for values spanning orders of magnitude. Returns path."""
    out = _out(out_path)
    validate_records(bars, "bars", required=("label", "value"), numeric=("value",),
                     example='{"label": "Invoices", "value": 10400000, "color": "#E8413B"}')
    validate_bounded_int(int(dpi), "dpi", min_val=36, max_val=600)
    labels = [str(b["label"]) for b in bars]
    vals = [float(b["value"]) for b in bars]
    cols = ["#" + validate_color(b.get("color", "#27B7E6")) for b in bars]
    fig, ax = plt.subplots(figsize=(9.0, 4.2), dpi=int(dpi))
    fig.patch.set_alpha(0)
    ax.set_facecolor("none")
    vlbl = {"path_effects": _STROKE, "color": "white", "fontsize": 13, "fontweight": "bold"}
    if horizontal:
        rects = ax.barh(range(len(bars)), vals, color=cols, height=0.66)
        ax.set_yticks(range(len(bars)))
        ax.set_yticklabels(labels, color="white", fontsize=15)
        ax.invert_yaxis()
        if log_scale:
            ax.set_xscale("log")
        ax.tick_params(axis="x", colors="#9fc4e8")
        if axis_label:
            ax.set_xlabel(axis_label, color="#9fc4e8", fontsize=12)
        for r, v in zip(rects, vals):
            ax.text(r.get_width() * (1.12 if log_scale else 1.0) + (0 if log_scale else max(vals) * 0.01),
                    r.get_y() + r.get_height() / 2, _fmt_value(v), va="center", **vlbl)
    else:
        rects = ax.bar(range(len(bars)), vals, color=cols, width=0.66)
        ax.set_xticks(range(len(bars)))
        ax.set_xticklabels(labels, color="white", fontsize=14)
        if log_scale:
            ax.set_yscale("log")
        ax.tick_params(axis="y", colors="#9fc4e8")
        if axis_label:
            ax.set_ylabel(axis_label, color="#9fc4e8", fontsize=12)
        for r, v in zip(rects, vals):
            ax.text(r.get_x() + r.get_width() / 2, r.get_height(), _fmt_value(v), ha="center", va="bottom", **vlbl)
    for s in ax.spines.values():
        s.set_visible(False)
    plt.tight_layout(pad=0.4)
    fig.savefig(str(out), transparent=True, bbox_inches="tight")
    plt.close(fig)
    log.info("bar_chart -> %s (%d bars)", out.name, len(bars))
    return str(out)


# --------------------------------------------------------------------------
# 4. Bubble cards
# --------------------------------------------------------------------------

def bubble_cards(out_path, cards, *, cols=2, width_px=2420, height_px=820, accent_stripe=True):
    """Render a grid of rounded "bubble" cards on transparency (supersampled).
    Each card = ``{"badge","title","subtitle","tag","color"}``. ``color`` tints
    the left accent stripe, the badge ring and the pill (pill text auto-darkens
    on light colours). Card heights are uniform and fill ``height_px``.
    Returns the output path."""
    out = _out(out_path)
    validate_records(cards, "cards", required=("title",),
                     example='{"badge": "C-5", "title": "Server root", "subtitle": "Full takeover", "tag": "CRITICAL", "color": "#C0202A"}')
    cols = validate_bounded_int(int(cols), "cols", min_val=1, max_val=6)
    _check_dims(width_px, height_px)
    W0, H0 = int(width_px), int(height_px)
    S = _ss_for(W0, H0)
    W, H = W0 * S, H0 * S
    rows = math.ceil(len(cards) / cols)
    gap = 24 * S
    cw = (W - (cols + 1) * gap) // cols
    ch = int((H - (rows + 1) * gap) / rows)
    rad = 26 * S
    base = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    shadow = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    sd = ImageDraw.Draw(shadow)
    for i in range(len(cards)):
        r, c = divmod(i, cols)
        x, y = gap + c * (cw + gap), gap + r * (ch + gap)
        sd.rounded_rectangle([x + 7 * S, y + 10 * S, x + cw + 7 * S, y + ch + 10 * S], radius=rad, fill=(0, 0, 0, 115))
    base = Image.alpha_composite(base, shadow.filter(ImageFilter.GaussianBlur(9 * S)))
    for i, card in enumerate(cards):
        r, c = divmod(i, cols)
        x, y = gap + c * (cw + gap), gap + r * (ch + gap)
        col = _rgb(card.get("color", "#27B7E6"))
        ptc = (12, 26, 47) if sum(col) / 3 > 150 else (255, 255, 255)
        stripe = int(ch * 0.05) if accent_stripe else 0
        _card(base, [x, y, x + cw, y + ch], rad, accent=col, stripe=stripe, ow=2 * S)
        d = ImageDraw.Draw(base)
        pad = int(ch * 0.16)
        right = x + cw - pad
        tag = card.get("tag")
        if tag:
            pf = _font(ch * 0.15)
            pb = d.textbbox((0, 0), str(tag), font=pf)
            pw, ph = (pb[2] - pb[0]) + 42 * S, int(ch * 0.30)
            px_, py_ = x + cw - pad - pw, y + (ch - ph) // 2
            d.rounded_rectangle([px_, py_, px_ + pw, py_ + ph], radius=ph // 2, fill=col)
            _gloss(base, px_ + pw / 2, py_ + ph / 2, ph / 2)
            d = ImageDraw.Draw(base)
            _ctext(d, px_ + pw / 2, py_ + ph / 2, tag, pf, ptc)
            right = px_
        tx = x + stripe + pad
        badge = card.get("badge")
        if badge:
            bd = int(ch * 0.46)
            bx, by = x + stripe + pad, y + (ch - bd) // 2
            d.ellipse([bx, by, bx + bd, by + bd], fill=(20, 48, 88, 255), outline=col, width=max(4, bd // 22))
            _gloss(base, bx + bd / 2, by + bd / 2, bd / 2)
            d = ImageDraw.Draw(base)
            _ctext(d, bx + bd / 2, by + bd / 2, badge, _font(bd * 0.30), (255, 255, 255))
            tx = bx + bd + pad
        tw = right - tx - pad
        tf, inf = _font(ch * 0.135), _font(ch * 0.118, bold=False)
        tl = _wrap(d, card.get("title", ""), tf, tw)
        il = _wrap(d, card.get("subtitle", ""), inf, tw) if card.get("subtitle") else []
        lh_t, lh_i = int(ch * 0.175), int(ch * 0.155)
        total = len(tl) * lh_t + (10 * S + len(il) * lh_i if il else 0)
        ty = y + (ch - total) // 2
        for ln in tl:
            d.text((tx, ty), ln, font=tf, fill=(255, 255, 255))
            ty += lh_t
        if il:
            ty += 10 * S
            for ln in il:
                d.text((tx, ty), ln, font=inf, fill=(193, 214, 240))
                ty += lh_i
    if S > 1:
        base = base.resize((W0, H0), Image.LANCZOS)
    base.save(str(out))
    log.info("bubble_cards -> %s (%d cards, ss=%d)", out.name, len(cards), S)
    return str(out)


# --------------------------------------------------------------------------
# 5. Node infographic
# --------------------------------------------------------------------------

def _dashed(d, p0, p1, fill, width, dash, gap):
    x0, y0 = p0
    x1, y1 = p1
    dist = math.hypot(x1 - x0, y1 - y0) or 1
    steps = int(dist / (dash + gap))
    for s in range(steps + 1):
        a = s * (dash + gap) / dist
        b = min(1, (s * (dash + gap) + dash) / dist)
        d.line([(x0 + (x1 - x0) * a, y0 + (y1 - y0) * a),
                (x0 + (x1 - x0) * b, y0 + (y1 - y0) * b)], fill=fill, width=width)


def node_infographic(out_path, nodes, *, width_px=2420, height_px=820):
    """Render glossy circular value nodes on a dashed zigzag connector
    (supersampled, transparent). ``nodes`` = ``[{"value","label","color"}, ...]``.
    Returns the output path."""
    out = _out(out_path)
    validate_records(nodes, "nodes", required=("value", "label"),
                     example='{"value": "26", "label": "CRITICAL", "color": "#C0202A"}')
    _check_dims(width_px, height_px)
    W0, H0 = int(width_px), int(height_px)
    S = _ss_for(W0, H0)
    W, H = W0 * S, H0 * S
    n = len(nodes)
    margin = int(W * 0.10)
    span = (W - 2 * margin) / max(1, n - 1) if n > 1 else 0
    r = int(min(H * 0.135, span * 0.28)) if n > 1 else int(H * 0.135)
    hi, lo = int(H * 0.40), int(H * 0.66)
    pts = [((int(margin + span * i) if n > 1 else W // 2), (hi if i % 2 else lo)) for i in range(n)]
    img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    for a, b in zip(pts, pts[1:]):
        _dashed(d, a, b, (110, 190, 255, 150), 7 * S, 26 * S, 18 * S)
    sh = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    sdd = ImageDraw.Draw(sh)
    for (x, y) in pts:
        sdd.ellipse([x - r + 6 * S, y - r + 10 * S, x + r + 6 * S, y + r + 10 * S], fill=(0, 0, 0, 120))
    img = Image.alpha_composite(img, sh.filter(ImageFilter.GaussianBlur(12 * S)))
    d = ImageDraw.Draw(img)
    for (x, y), nd in zip(pts, nodes):
        col = _rgb(nd.get("color", "#27B7E6"))
        dark = tuple(int(c * 0.62) for c in col)
        d.ellipse([x - r - 8 * S, y - r - 8 * S, x + r + 8 * S, y + r + 8 * S], outline=(*col, 90), width=4 * S)
        # radial-ish: paste a vertical gradient (lighter top -> darker bottom) masked to a circle
        grad = _vgrad(2 * r, 2 * r, col, dark, 255)
        m = Image.new("L", (2 * r, 2 * r), 0)
        ImageDraw.Draw(m).ellipse([0, 0, 2 * r - 1, 2 * r - 1], fill=255)
        img.paste(grad, (x - r, y - r), m)
        _gloss(img, x, y, r)
        d = ImageDraw.Draw(img)
        _ctext(d, x, y - r * 0.05, nd.get("value", ""), _font(r * 0.78), (255, 255, 255))
        ly = y + r + int(H * 0.04)
        for j, ln in enumerate(str(nd.get("label", "")).split("\n")):
            _ctext(d, x, ly + j * int(H * 0.046), ln, _font(H * 0.039), (235, 244, 255))
    if S > 1:
        img = img.resize((W0, H0), Image.LANCZOS)
    img.save(str(out))
    log.info("node_infographic -> %s (%d nodes, ss=%d)", out.name, n, S)
    return str(out)


# --------------------------------------------------------------------------
# 6. Numbered cards
# --------------------------------------------------------------------------

def numbered_cards(out_path, items, *, cols=2, width_px=2420, height_px=940):
    """Render numbered cards (recommendations / steps) in columns (supersampled).
    ``items`` = ``[{"title","subtitle","color"}, ...]`` auto-numbered 1..N
    column-major; ``color`` tints the gradient number badge. Returns path."""
    out = _out(out_path)
    validate_records(items, "items", required=("title",),
                     example='{"title": "Contain the intrusion", "subtitle": "Kill C2 beacons", "color": "#C0202A"}')
    cols = validate_bounded_int(int(cols), "cols", min_val=1, max_val=6)
    _check_dims(width_px, height_px)
    W0, H0 = int(width_px), int(height_px)
    S = _ss_for(W0, H0)
    W, H = W0 * S, H0 * S
    per = math.ceil(len(items) / cols)
    gap = 26 * S
    cw = (W - (cols + 1) * gap) // cols
    ch = int((H - (per + 1) * gap) / per)
    rad = 24 * S
    base = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    shadow = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    sd = ImageDraw.Draw(shadow)
    for i in range(len(items)):
        c, r = i // per, i % per
        x, y = gap + c * (cw + gap), gap + r * (ch + gap)
        sd.rounded_rectangle([x + 6 * S, y + 9 * S, x + cw + 6 * S, y + ch + 9 * S], radius=rad, fill=(0, 0, 0, 110))
    base = Image.alpha_composite(base, shadow.filter(ImageFilter.GaussianBlur(9 * S)))
    for i, it in enumerate(items):
        c, r = i // per, i % per
        x, y = gap + c * (cw + gap), gap + r * (ch + gap)
        col = _rgb(it.get("color", "#27B7E6"))
        _card(base, [x, y, x + cw, y + ch], rad, ow=2 * S)
        d = ImageDraw.Draw(base)
        pad = int(ch * 0.16)
        bd = int(ch * 0.42)
        bx, by = x + pad, y + (ch - bd) // 2
        dark = tuple(int(cc * 0.62) for cc in col)
        grad = _vgrad(bd, bd, col, dark, 255)
        m = Image.new("L", (bd, bd), 0)
        ImageDraw.Draw(m).ellipse([0, 0, bd - 1, bd - 1], fill=255)
        base.paste(grad, (bx, by), m)
        d.ellipse([bx, by, bx + bd, by + bd], outline=(255, 255, 255, 70), width=3 * S)
        _gloss(base, bx + bd / 2, by + bd / 2, bd / 2)
        d = ImageDraw.Draw(base)
        _ctext(d, bx + bd / 2, by + bd / 2, str(it.get("n", i + 1)), _font(bd * 0.42), (255, 255, 255))
        tx = bx + bd + pad
        tw = x + cw - tx - pad
        tf, af = _font(ch * 0.155), _font(ch * 0.125, bold=False)
        tl = _wrap(d, it.get("title", ""), tf, tw)
        al = _wrap(d, it.get("subtitle", ""), af, tw) if it.get("subtitle") else []
        lh_t, lh_a = int(ch * 0.195), int(ch * 0.165)
        total = len(tl) * lh_t + (8 * S + len(al) * lh_a if al else 0)
        ty = y + (ch - total) // 2
        for ln in tl:
            d.text((tx, ty), ln, font=tf, fill=(255, 255, 255))
            ty += lh_t
        if al:
            ty += 8 * S
            for ln in al:
                d.text((tx, ty), ln, font=af, fill=(179, 203, 233))
                ty += lh_a
    if S > 1:
        base = base.resize((W0, H0), Image.LANCZOS)
    base.save(str(out))
    log.info("numbered_cards -> %s (%d items, ss=%d)", out.name, len(items), S)
    return str(out)


# --------------------------------------------------------------------------
# 7. Decorative panel
# --------------------------------------------------------------------------

def decorative_panel(out_path, *, width_px=1500, height_px=760, color="#5aa0e0", motif="lock"):
    """Render an abstract hexagon-cluster panel with an optional centre motif
    (``"lock"``, ``"shield"`` or ``"none"``) on transparency (supersampled) —
    a tasteful visual filler for title/closing slides. Returns the path."""
    out = _out(out_path)
    motif = validate_choice(motif, "motif", ("lock", "shield", "none")) or "none"
    _check_dims(width_px, height_px)
    W0, H0 = int(width_px), int(height_px)
    S = _ss_for(W0, H0)
    W, H = W0 * S, H0 * S
    col = _rgb(color)
    rng = random.Random(7)
    panel = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    pd = ImageDraw.Draw(panel)
    for _ in range(60):
        pd.polygon(_hexagon(rng.randint(0, W), rng.randint(0, H), rng.randint(40 * S, 130 * S)),
                   outline=(*col, rng.randint(12, 55)), width=2 * S)
    cx, cy = W // 2, H // 2
    for k, rr in enumerate([int(H * 0.31), int(H * 0.26), int(H * 0.21)]):
        pd.polygon(_hexagon(cx, cy, rr), outline=(80, 180, 255, 120 - k * 25), width=4 * S)
    light = (150, 210, 255, 235)
    if motif == "lock":
        pd.rounded_rectangle([cx - 70 * S, cy - 10 * S, cx + 70 * S, cy + 110 * S], radius=14 * S, outline=light, width=6 * S)
        pd.arc([cx - 45 * S, cy - 95 * S, cx + 45 * S, cy + 25 * S], start=180, end=360, fill=light, width=6 * S)
        pd.ellipse([cx - 14 * S, cy + 30 * S, cx + 14 * S, cy + 58 * S], fill=light)
    elif motif == "shield":
        pts = [(cx, cy - 110 * S), (cx + 80 * S, cy - 70 * S), (cx + 80 * S, cy + 20 * S),
               (cx, cy + 110 * S), (cx - 80 * S, cy + 20 * S), (cx - 80 * S, cy - 70 * S)]
        pd.polygon(pts, outline=light, width=6 * S)
        pd.line([(cx - 28 * S, cy), (cx - 6 * S, cy + 30 * S), (cx + 34 * S, cy - 28 * S)], fill=light, width=8 * S, joint="curve")
    panel = panel.filter(ImageFilter.GaussianBlur(0.6 * S))
    if S > 1:
        panel = panel.resize((W0, H0), Image.LANCZOS)
    panel.save(str(out))
    log.info("decorative_panel -> %s (%s, %dx%d, ss=%d)", out.name, motif, W0, H0, S)
    return str(out)
