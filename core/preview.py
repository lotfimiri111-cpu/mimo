"""
Preview Engine — مذكرتي Pro v18 Commercial
PPTX → PDF → JPG images with watermark protection.
Images are cached per order for performance.
"""
from __future__ import annotations

import io
import logging
import os
import subprocess
import tempfile
import time
from pathlib import Path
from typing import Optional

from PIL import Image, ImageDraw, ImageFont

log = logging.getLogger(__name__)

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CACHE_DIR = os.path.join(BASE_DIR, "storage", "preview_cache")
STORAGE_DIR = os.path.join(BASE_DIR, "storage")


def _ensure_dirs():
    os.makedirs(CACHE_DIR, exist_ok=True)


def _slide_cache_path(order_id: str, slide_n: int) -> str:
    order_dir = os.path.join(CACHE_DIR, order_id)
    os.makedirs(order_dir, exist_ok=True)
    return os.path.join(order_dir, f"slide_{slide_n:02d}.jpg")


def _thumb_cache_path(order_id: str, slide_n: int) -> str:
    order_dir = os.path.join(CACHE_DIR, order_id)
    os.makedirs(order_dir, exist_ok=True)
    return os.path.join(order_dir, f"thumb_{slide_n:02d}.jpg")


def _pdf_path(order_id: str) -> str:
    return os.path.join(CACHE_DIR, order_id, "slides.pdf")


def pptx_to_pdf(pptx_path: str, order_id: str) -> Optional[str]:
    """Convert PPTX to PDF using LibreOffice. Returns PDF path or None."""
    _ensure_dirs()
    order_dir = os.path.join(CACHE_DIR, order_id)
    os.makedirs(order_dir, exist_ok=True)
    out_pdf = _pdf_path(order_id)

    if os.path.exists(out_pdf):
        return out_pdf  # cached

    try:
        t0 = time.monotonic()
        r = subprocess.run(
            ["libreoffice", "--headless", "--convert-to", "pdf",
             "--outdir", order_dir, pptx_path],
            capture_output=True, text=True, timeout=120
        )
        if r.returncode != 0:
            log.error(f"LibreOffice failed: {r.stderr[:300]}")
            return None

        # LibreOffice names it after the source file — rename to slides.pdf
        pptx_name = os.path.splitext(os.path.basename(pptx_path))[0]
        generated_pdf = os.path.join(order_dir, pptx_name + ".pdf")
        if os.path.exists(generated_pdf) and generated_pdf != out_pdf:
            os.rename(generated_pdf, out_pdf)

        if not os.path.exists(out_pdf):
            log.error("PDF not created after conversion")
            return None

        log.info(f"PDF created: {out_pdf} ({os.path.getsize(out_pdf):,}B) in {time.monotonic()-t0:.1f}s")
        return out_pdf
    except Exception as e:
        log.error(f"pptx_to_pdf error: {e}", exc_info=True)
        return None


def _add_watermark(img: Image.Image, text: str = "مذكرتي Pro — للمعاينة فقط") -> Image.Image:
    """Add a subtle diagonal watermark."""
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)

    # Use a built-in font (no external font needed)
    try:
        font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 28)
    except Exception:
        font = ImageFont.load_default()

    w, h = img.size
    # Draw watermark pattern diagonally
    for x in range(-w, w * 2, 320):
        for y in range(-h, h * 2, 180):
            draw.text((x, y), text, fill=(255, 255, 255, 28), font=font)

    # Convert base image to RGBA for compositing
    if img.mode != "RGBA":
        img = img.convert("RGBA")
    watermarked = Image.alpha_composite(img, overlay)
    return watermarked.convert("RGB")


def render_slide(pptx_path: str, order_id: str, slide_n: int,
                 watermark: bool = True) -> Optional[bytes]:
    """
    Render slide N (1-based) as JPEG bytes.
    Returns None if not available.
    """
    _ensure_dirs()
    cache_path = _slide_cache_path(order_id, slide_n)

    # Return from cache if exists
    if os.path.exists(cache_path):
        with open(cache_path, "rb") as f:
            return f.read()

    # Need PDF first
    pdf_path = _pdf_path(order_id)
    if not os.path.exists(pdf_path):
        pdf_path = pptx_to_pdf(pptx_path, order_id)
        if not pdf_path:
            return None

    try:
        from pdf2image import convert_from_path
        # Render only the specific page (1-based in pdf2image)
        pages = convert_from_path(
            pdf_path,
            dpi=150,
            first_page=slide_n,
            last_page=slide_n,
            fmt="jpeg",
            thread_count=1,
        )
        if not pages:
            return None

        img = pages[0]
        if watermark:
            img = _add_watermark(img)

        # Save to cache
        img.save(cache_path, "JPEG", quality=88, optimize=True)

        with open(cache_path, "rb") as f:
            return f.read()
    except Exception as e:
        log.error(f"render_slide error: slide={slide_n} {e}", exc_info=True)
        return None


def render_thumbnail(pptx_path: str, order_id: str, slide_n: int) -> Optional[bytes]:
    """Render small thumbnail (300px wide) for the sidebar."""
    _ensure_dirs()
    cache_path = _thumb_cache_path(order_id, slide_n)

    if os.path.exists(cache_path):
        with open(cache_path, "rb") as f:
            return f.read()

    # Get full slide first
    full_bytes = render_slide(pptx_path, order_id, slide_n, watermark=False)
    if not full_bytes:
        return None

    try:
        img = Image.open(io.BytesIO(full_bytes))
        # Resize to thumbnail
        w, h = img.size
        thumb_w = 300
        thumb_h = int(h * thumb_w / w)
        thumb = img.resize((thumb_w, thumb_h), Image.LANCZOS)

        buf = io.BytesIO()
        thumb.save(buf, "JPEG", quality=75)
        data = buf.getvalue()

        with open(cache_path, "wb") as f:
            f.write(data)
        return data
    except Exception as e:
        log.error(f"thumbnail error: {e}")
        return None


def get_slide_count_from_pdf(order_id: str) -> int:
    """Get number of slides from cached PDF."""
    pdf_path = _pdf_path(order_id)
    if not os.path.exists(pdf_path):
        return 0
    try:
        from pdf2image import pdfinfo_from_path
        info = pdfinfo_from_path(pdf_path)
        return info.get("Pages", 0)
    except Exception:
        return 0


def prerender_all(pptx_path: str, order_id: str, slide_count: int):
    """Background pre-render all slides for fast access."""
    try:
        pdf_path = pptx_to_pdf(pptx_path, order_id)
        if not pdf_path:
            return
        log.info(f"Pre-rendering {slide_count} slides for order {order_id[:8]}...")
        for n in range(1, slide_count + 1):
            render_thumbnail(pptx_path, order_id, n)
        log.info(f"Pre-render complete for {order_id[:8]}")
    except Exception as e:
        log.error(f"prerender_all error: {e}")


def clear_cache(order_id: str):
    """Delete cached preview images for an order."""
    import shutil
    order_dir = os.path.join(CACHE_DIR, order_id)
    if os.path.exists(order_dir):
        shutil.rmtree(order_dir)
