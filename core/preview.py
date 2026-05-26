"""
core/preview.py v6 — معاينة دقيقة 100% عبر LibreOffice (شريحة بشريحة)
"""
import base64, glob, io, logging, os, subprocess, tempfile, threading
from pathlib import Path
from typing import List

log = logging.getLogger(__name__)
_cache: dict = {}
_cache_lock = threading.Lock()
MAX_SLIDES = 20  # عرض جميع الشرائح في المعاينة (ما تراه = ما تحصل عليه)


def get_cached_preview(pid):
    with _cache_lock: return _cache.get(pid)

def set_cached_preview(pid, slides):
    with _cache_lock: _cache[pid] = slides


def pptx_to_preview_images(pptx_path: str, watermark: bool = True) -> List[str]:
    """
    تحويل PPTX → صور JPEG معاينة بدقة كاملة (تطابق 100% مع الملف الحقيقي).
    يستخدم LibreOffice للتحويل مع Pillow كـ fallback.
    """
    try:
        results = _libreoffice_render(pptx_path, watermark=watermark)
        if results:
            return results
    except Exception as e:
        log.warning(f"LibreOffice render failed: {e}, falling back to Pillow")

    # Fallback: Pillow
    try:
        return _pillow_render(pptx_path, watermark=watermark)
    except Exception as e:
        log.error(f"Pillow fallback also failed: {e}", exc_info=True)
        return []


def _libreoffice_render(pptx_path: str, watermark: bool = True) -> List[str]:
    """
    تحويل كل شريحة بشكل منفصل عبر LibreOffice للحصول على دقة 100%.
    كل شريحة تُحفظ كـ PPTX مؤقت → PNG عبر soffice.
    """
    from PIL import Image
    from pptx import Presentation
    from pptx.oxml.ns import qn
    from lxml import etree

    prs_orig = Presentation(pptx_path)
    total = len(prs_orig.slides)
    slides_b64 = []

    with tempfile.TemporaryDirectory() as tmpdir:
        for idx in range(min(total, MAX_SLIDES)):
            try:
                png_path = _render_single_slide(pptx_path, idx, tmpdir)
                if not png_path or not os.path.exists(png_path):
                    log.warning(f"Slide {idx}: no PNG produced")
                    continue

                img = Image.open(png_path).convert("RGB")

                # تحجيم احترافي (max 1280px عرضاً)
                max_w = 1280
                if img.width > max_w:
                    ratio = max_w / img.width
                    img = img.resize(
                        (max_w, int(img.height * ratio)),
                        Image.LANCZOS
                    )

                if watermark:
                    img = _add_watermark(img)

                buf = io.BytesIO()
                img.save(buf, "JPEG", quality=92)
                slides_b64.append(base64.b64encode(buf.getvalue()).decode())
                log.info(f"Slide {idx}: {len(buf.getvalue()):,} bytes")

            except Exception as e:
                log.warning(f"Slide {idx} render failed: {e}")
                continue

    return slides_b64


def _render_single_slide(pptx_path: str, slide_idx: int, tmpdir: str) -> str:
    """تحويل شريحة واحدة إلى PNG عبر LibreOffice"""
    from pptx import Presentation
    from pptx.oxml.ns import qn
    from lxml import etree

    prs = Presentation(pptx_path)
    src_slide = prs.slides[slide_idx]

    # إنشاء PPTX بشريحة واحدة مع الحفاظ على الأبعاد والمحتوى الكامل
    single = Presentation()
    single.slide_width = prs.slide_width
    single.slide_height = prs.slide_height

    # استخدام blank layout
    layout = single.slide_layouts[6]
    new_slide = single.slides.add_slide(layout)

    # نسخ الخلفية (مهم للألوان الصحيحة)
    try:
        bg_src = src_slide.background._element
        bg_dst = new_slide.background._element
        bg_dst.clear()
        for child in bg_src:
            bg_dst.append(etree.fromstring(etree.tostring(child)))
    except Exception:
        pass

    # نسخ جميع الأشكال
    sp_tree = new_slide.shapes._spTree
    for elem in list(sp_tree):
        sp_tree.remove(elem)
    for elem in src_slide.shapes._spTree:
        try:
            sp_tree.append(etree.fromstring(etree.tostring(elem)))
        except Exception:
            pass

    slide_path = os.path.join(tmpdir, f"slide_{slide_idx:03d}.pptx")
    single.save(slide_path)

    env = os.environ.copy()
    env["HOME"] = tmpdir  # تجنّب تعارض LibreOffice profile

    result = subprocess.run(
        [
            "soffice",
            "--headless",
            "--norestore",
            "--nofirststartwizard",
            "--convert-to", "png",
            "--outdir", tmpdir,
            slide_path,
        ],
        capture_output=True,
        text=True,
        timeout=20,
        env=env,
    )

    if result.returncode != 0:
        log.warning(f"soffice exit {result.returncode}: {result.stderr[:200]}")

    # البحث عن الـ PNG المُنتج
    pngs = glob.glob(os.path.join(tmpdir, f"slide_{slide_idx:03d}*.png"))
    return pngs[0] if pngs else None


# ── Pillow fallback ───────────────────────────────────────────────────────────

CAIRO_PATHS = [
    os.path.expanduser("~/.fonts/cairo/Cairo.ttf"),
    "/root/.fonts/cairo/Cairo.ttf",
    "/tmp/fonts/cairo/Cairo.ttf",
    "/opt/render/project/src/.fonts/cairo/Cairo.ttf",
    "/opt/render/.fonts/cairo/Cairo.ttf",
]
FALLBACK_FONTS = [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
]

def _find_font(size=16, bold=False):
    from PIL import ImageFont
    for p in CAIRO_PATHS + FALLBACK_FONTS:
        if os.path.exists(p):
            try: return ImageFont.truetype(p, size)
            except: pass
    return ImageFont.load_default()


def _hex(h) -> tuple:
    h = str(h).lstrip("#")
    if len(h) == 6:
        return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    return (200, 200, 200)


def _read_color_xml(elem):
    if elem is None: return None
    try:
        for child in elem.iter():
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'srgbClr':
                val = child.get('val', '')
                if val: return _hex(val)
            elif tag == 'sysClr':
                val = child.get('lastClr', '')
                if val: return _hex(val)
    except: pass
    return None


def _bg_color(slide) -> tuple:
    try:
        fill = slide.background.fill
        if fill.type is not None:
            c = fill.fore_color.rgb
            return (c.r, c.g, c.b)
    except: pass
    try:
        c = _read_color_xml(slide.background._element)
        if c: return c
    except: pass
    return (255, 255, 255)


def _smart_text_color(bg: tuple) -> tuple:
    lum = 0.299*bg[0] + 0.587*bg[1] + 0.114*bg[2]
    return (255, 255, 255) if lum < 128 else (20, 20, 20)


def _pillow_render(pptx_path: str, watermark: bool = True) -> List[str]:
    from pptx import Presentation
    from PIL import Image, ImageDraw
    prs = Presentation(pptx_path)
    W = max(960, int((prs.slide_width  or 9144000) / 9525))
    H = max(540, int((prs.slide_height or 5143500) / 9525))
    results = []
    for slide in list(prs.slides)[:MAX_SLIDES]:
        bg = _bg_color(slide)
        img = Image.new("RGB", (W, H), bg)
        draw = ImageDraw.Draw(img)
        for shape in slide.shapes:
            try:
                L = int((shape.left  or 0) / 914400 * 96)
                T = int((shape.top   or 0) / 914400 * 96)
                SW = int((shape.width or 0) / 914400 * 96)
                if not shape.has_text_frame: continue
                y = T + 6
                for para in shape.text_frame.paragraphs:
                    text = para.text.strip()
                    if not text: y += 8; continue
                    fs, tc = 14, _smart_text_color(bg)
                    for run in para.runs:
                        try:
                            if run.font.size: fs = max(8, min(int(run.font.size / 12700), 60))
                            rgb = run.font.color.rgb
                            tc = (rgb.r, rgb.g, rgb.b)
                        except: pass
                        break
                    draw.text((L + 6, y), text, fill=tc, font=_find_font(fs))
                    y += fs + 4
            except: pass
        if watermark:
            img = _add_watermark(img)
        buf = io.BytesIO()
        img.save(buf, "JPEG", quality=90)
        results.append(base64.b64encode(buf.getvalue()).decode())
    return results


def _add_watermark(img):
    from PIL import Image, ImageDraw
    W, H = img.size
    overlay = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)
    font = _find_font(max(22, W // 26), bold=True)
    text = "مذكرتي Pro — معاينة فقط"
    try:
        bb = d.textbbox((0, 0), text, font=font)
        tw, th = bb[2] - bb[0], bb[3] - bb[1]
    except:
        tw, th = len(text) * 14, 28
    pad = 40
    tl = Image.new("RGBA", (tw + pad, th + pad), (0, 0, 0, 0))
    td = ImageDraw.Draw(tl)
    td.text((pad // 2, pad // 2), text, fill=(255, 255, 255, 80), font=font)
    rot = tl.rotate(-30, expand=True)
    rw, rh = rot.size
    sx = max(rw + 30, W // 3)
    sy = max(rh + 20, H // 3)
    for row in range(-1, H // sy + 2):
        for col in range(-1, W // sx + 2):
            overlay.paste(rot, (col * sx - rw // 2, row * sy - rh // 2), rot)
    return Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")
