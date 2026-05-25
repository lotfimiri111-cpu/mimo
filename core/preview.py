"""
core/preview.py v3 — معاينة PPTX حقيقية بدون LibreOffice
يستخدم python-pptx لقراءة كل عناصر الشريحة ويحوّلها إلى صورة JPEG دقيقة بـ Pillow.
"""
import base64
import io
import logging
import threading
from pathlib import Path
from typing import List, Optional, Tuple

log = logging.getLogger(__name__)

_cache: dict = {}
_cache_lock = threading.Lock()
MAX_PREVIEW_SLIDES = 3


def get_cached_preview(presentation_id: str) -> Optional[List[str]]:
    with _cache_lock:
        return _cache.get(presentation_id)


def set_cached_preview(presentation_id: str, slides: List[str]) -> None:
    with _cache_lock:
        _cache[presentation_id] = slides


def pptx_to_preview_images(pptx_path: str, watermark: bool = True) -> List[str]:
    try:
        slides = _render_pptx(pptx_path)
        if watermark:
            slides = [_add_watermark(s) for s in slides]
        return slides
    except Exception as exc:
        log.warning(f"Preview render failed: {exc}", exc_info=True)
        return []


# ─── EMU helpers ──────────────────────────────────────────────────────────────
EMU = 914400.0
SCALE = 96 / EMU   # EMU → px at 96dpi


def _px(emu) -> int:
    return int((emu or 0) * SCALE)


def _rgb_from_pptx(color_obj) -> Optional[Tuple[int,int,int]]:
    try:
        rgb = color_obj.rgb
        return (rgb.r, rgb.g, rgb.b)
    except Exception:
        return None


def _hex_to_rgb(h: str) -> Tuple[int,int,int]:
    h = h.lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))


# ─── Main renderer ────────────────────────────────────────────────────────────
def _render_pptx(pptx_path: str) -> List[str]:
    from pptx import Presentation
    from pptx.enum.dml import MSO_THEME_COLOR
    from PIL import Image

    prs = Presentation(pptx_path)
    W = max(960, _px(prs.slide_width))
    H = max(540, _px(prs.slide_height))

    results = []
    for slide in list(prs.slides)[:MAX_PREVIEW_SLIDES]:
        img = _draw_slide(slide, W, H)
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=88, optimize=True)
        results.append(base64.b64encode(buf.getvalue()).decode("ascii"))
    return results


def _draw_slide(slide, W: int, H: int):
    from PIL import Image, ImageDraw
    from pptx.util import Pt
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    # 1) خلفية الشريحة
    img = _draw_background(slide, W, H)
    draw = ImageDraw.Draw(img)

    # 2) ارسم كل شكل بالترتيب
    for shape in slide.shapes:
        try:
            _draw_shape(img, draw, shape, W, H)
        except Exception as e:
            log.debug(f"Shape render skip: {e}")

    return img


def _draw_background(slide, W: int, H: int):
    from PIL import Image, ImageDraw
    from pptx.dml.color import RGBColor

    bg_color = (255, 255, 255)
    try:
        bg = slide.background
        fill = bg.fill
        ftype = fill.type
        if ftype is not None:
            c = _rgb_from_pptx(fill.fore_color)
            if c:
                bg_color = c
    except Exception:
        pass

    img = Image.new("RGB", (W, H), bg_color)

    # gradient إذا كان محدداً (نحاكيه بشكل بسيط)
    # نقرأ الـ XML للحصول على gradient stops
    try:
        from pptx.oxml.ns import qn
        from lxml import etree
        import re

        bg_elem = slide.background._element
        grad_fill = bg_elem.find('.//' + qn('a:gradFill'))
        if grad_fill is not None:
            stops = []
            for gs in grad_fill.findall('.//' + qn('a:gs')):
                pos = int(gs.get('pos', '0')) / 100000
                srgb = gs.find('.//' + qn('a:srgbClr'))
                if srgb is not None:
                    val = srgb.get('val', 'FFFFFF')
                    stops.append((pos, _hex_to_rgb(val)))
            if len(stops) >= 2:
                img = _draw_gradient(W, H, stops)
    except Exception:
        pass

    return img


def _draw_gradient(W: int, H: int, stops: list):
    """يرسم gradient عمودي بين عدة ألوان."""
    from PIL import Image
    import numpy as np

    arr = np.zeros((H, W, 3), dtype=np.uint8)
    stops_sorted = sorted(stops, key=lambda x: x[0])

    for y in range(H):
        t = y / H
        # إيجاد الـ stop المناسب
        c1 = stops_sorted[0][1]
        c2 = stops_sorted[-1][1]
        for i in range(len(stops_sorted)-1):
            p0, col0 = stops_sorted[i]
            p1, col1 = stops_sorted[i+1]
            if p0 <= t <= p1:
                if p1 - p0 < 0.001:
                    local = 0
                else:
                    local = (t - p0) / (p1 - p0)
                c1 = tuple(int(col0[j] + (col1[j] - col0[j]) * local) for j in range(3))
                c2 = c1
                break
        arr[y, :] = c1

    return Image.fromarray(arr, "RGB")


def _draw_shape(img, draw, shape, W: int, H: int):
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    from PIL import Image

    left   = _px(shape.left)
    top    = _px(shape.top)
    width  = _px(shape.width)
    height = _px(shape.height)

    # ── رسم خلفية الشكل ──────────────────────────────────────────────
    try:
        fill = shape.fill
        ftype = fill.type
        if ftype is not None and width > 0 and height > 0:
            fc = _rgb_from_pptx(fill.fore_color)
            if fc:
                # شكل مستطيل أو مدوّر
                _draw_rect_shape(img, draw, shape, left, top, width, height, fc)
    except Exception:
        pass

    # ── رسم النص ─────────────────────────────────────────────────────
    if shape.has_text_frame:
        _draw_text_frame(img, draw, shape, left, top, width, height)


def _draw_rect_shape(img, draw, shape, left, top, width, height, fill_color):
    from PIL import Image, ImageDraw
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    # تحقق نوع الشكل للحواف المدوّرة
    try:
        shape_type = shape.shape_type
        # oval/ellipse
        if shape_type == MSO_SHAPE_TYPE.FREEFORM or \
           (hasattr(shape, 'auto_shape_type') and 'OVAL' in str(shape.auto_shape_type)):
            draw.ellipse([left, top, left+width, top+height], fill=fill_color)
            return
    except Exception:
        pass

    draw.rectangle([left, top, left+width, top+height], fill=fill_color)


def _draw_text_frame(img, draw, shape, left, top, width, height):
    from PIL import ImageFont
    from pptx.enum.text import PP_ALIGN

    tf = shape.text_frame
    y_cursor = top + 6

    for para in tf.paragraphs:
        if not para.text.strip():
            y_cursor += 8
            continue

        # حجم الخط ولونه من أول run
        fsize = 14
        fcolor = (30, 30, 30)
        bold = False
        for run in para.runs:
            try:
                if run.font.size:
                    fsize = max(7, min(int(run.font.size / 12700), 72))
                if run.font.bold:
                    bold = True
                c = _rgb_from_pptx(run.font.color)
                if c:
                    fcolor = c
            except Exception:
                pass
            break

        # اختر خط مناسب
        font = _get_font(fsize, bold)

        text = para.text.strip()
        if not text:
            continue

        # محاذاة
        try:
            align = para.alignment
        except Exception:
            align = None

        # اقتصاص النص إذا كان طويلاً
        text = _wrap_text(text, font, width - 12)

        # x بحسب المحاذاة
        x = left + 6
        if align == PP_ALIGN.CENTER:
            x = left + width // 2 - _text_width(text.split('\n')[0], font) // 2
        elif align == PP_ALIGN.RIGHT or align is None:
            # العربية RTL — نضع النص على اليمين
            x = left + width - _text_width(text.split('\n')[0], font) - 6

        draw.text((x, y_cursor), text, fill=fcolor, font=font)
        lines = text.count('\n') + 1
        y_cursor += (fsize + 4) * lines

        if y_cursor > top + height:
            break


def _get_font(size: int, bold: bool = False):
    from PIL import ImageFont

    paths = [
        "/home/user/.fonts/cairo/Cairo.ttf",
        "/root/.fonts/cairo/Cairo.ttf",
        "/usr/share/fonts/truetype/cairo/Cairo.ttf",
        "/tmp/fonts/cairo/Cairo.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for p in paths:
        try:
            return ImageFont.truetype(p, size)
        except Exception:
            continue
    return ImageFont.load_default()


def _text_width(text: str, font) -> int:
    try:
        from PIL import Image, ImageDraw
        tmp = ImageDraw.Draw(Image.new("RGB", (1, 1)))
        bbox = tmp.textbbox((0, 0), text, font=font)
        return bbox[2] - bbox[0]
    except Exception:
        return len(text) * 8


def _wrap_text(text: str, font, max_width: int) -> str:
    if max_width <= 0:
        return text
    words = text.split()
    lines = []
    current = ""
    for word in words:
        test = (current + " " + word).strip()
        if _text_width(test, font) <= max_width:
            current = test
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return "\n".join(lines) if lines else text


# ─── Watermark ────────────────────────────────────────────────────────────────
def _add_watermark(b64_jpeg: str) -> str:
    import os
    from PIL import Image, ImageDraw, ImageFont

    data = base64.b64decode(b64_jpeg)
    img = Image.open(io.BytesIO(data)).convert("RGBA")
    W, H = img.size

    overlay = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)

    font_size = max(24, W // 24)
    font = _get_font(font_size, bold=True)
    text = "مذكرتي Pro — معاينة فقط"

    # قياس النص
    try:
        tmp_draw = ImageDraw.Draw(Image.new("RGBA", (1,1)))
        bbox = tmp_draw.textbbox((0,0), text, font=font)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]
    except Exception:
        tw, th = font_size * 15, font_size + 4

    # إنشاء طبقة للعلامة المائية المدوّرة
    txt_layer = Image.new("RGBA", (tw + 40, th + 20), (0,0,0,0))
    td = ImageDraw.Draw(txt_layer)
    td.text((20, 10), text, fill=(255,255,255,80), font=font)
    rotated = txt_layer.rotate(-30, expand=True)

    # توزيع العلامة المائية بشكل شبكي
    step_x = max(rotated.width + 40, W // 3)
    step_y = max(rotated.height + 30, H // 3)
    for row in range(-1, H // step_y + 2):
        for col in range(-1, W // step_x + 2):
            x = col * step_x - rotated.width // 2
            y = row * step_y - rotated.height // 2
            overlay.paste(rotated, (x, y), rotated)

    combined = Image.alpha_composite(img, overlay).convert("RGB")
    buf = io.BytesIO()
    combined.save(buf, format="JPEG", quality=85, optimize=True)
    return base64.b64encode(buf.getvalue()).decode("ascii")
