"""
Classic Engine Slides — مذكرتي Pro v17.3
تخطيط أكاديمي كلاسيكي: هيدر علوي + محتوى يملأ الشريحة 100%
"""
from __future__ import annotations
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from engine.primitives import (
    W, H, rect, rrect, oval, bg, hline, vline,
    gradient_fill, gradient_rect, shadow, set_solid_alpha,
    multi_stop_gradient, diamond, decorative_dots,
    slide_number, txt, blank_slide,
)
from core.themes import Theme
from core.models import PresentationRequest

_FONT = "Cairo"

HEADER_H = 2.55     # هيدر أكبر يملأ الثلث العلوي
FOOTER_H = 0.32     # شريط سفلي رفيع
MX = 1.2            # هامش أفقي أضيق لمساحة أكبر


def set_font(font_name: str):
    global _FONT
    _FONT = font_name


def _header(slide, T: Theme, title: str, page_num: int = 0, req=None):
    bg(slide, T.bg_rgb)
    # تدرج خلفية الهيدر
    gradient_rect(slide, 0, 0, W, HEADER_H, T.grad2, T.grad1, angle=90)
    # خط accent سميك أسفل الهيدر
    al = rect(slide, 0, HEADER_H-0.18, W, 0.18, T.accent_rgb)
    if al: multi_stop_gradient(al,[(0,T.bg),(40,T.accent),(60,T.accent2),(100,T.bg)],0)
    # خط رفيع فوقه
    rect(slide, 0, HEADER_H-0.28, W, 0.06, T.muted_rgb)
    # شريط acc يميني
    vr = rect(slide, W-0.45, 0, 0.45, HEADER_H, T.accent_rgb)
    if vr: gradient_fill(vr, T.accent_grad1, T.accent_grad2, 90)
    # زخرفة
    oval(slide, W-5, -2, 7, 7, T.accent_rgb, alpha=8)
    decorative_dots(slide, MX, HEADER_H*0.15, 4, 2, 0.14, 0.32, T.accent_rgb, alpha=10)
    # عنوان — يبدأ بعد مساحة رقم الشريحة
    title_x = MX + (1.3 if page_num > 0 else 0)
    title_w = W - title_x - 0.55
    txt(slide, title, title_x, 0.22, title_w, HEADER_H-0.45,
        font=_FONT, size=22, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)
    # رقم الصفحة — دائرة واحدة فقط في الزاوية اليسرى
    if page_num > 0:
        nb_s = 0.78
        nb_x = 0.28
        nb_y = (HEADER_H - nb_s) / 2
        pb = rrect(slide, nb_x, nb_y, nb_s, nb_s, T.accent_rgb, radius_pct=50)
        if pb: multi_stop_gradient(pb,[(0,T.accent),(100,T.accent2)],135)
        txt(slide, str(page_num), nb_x, nb_y, nb_s, nb_s,
            font="Calibri", size=15, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)
        txt(slide, f"/{getattr(req, "_total_slides", 13)}", nb_x + nb_s, nb_y + nb_s*0.35, 0.8, nb_s*0.4,
            font="Calibri", size=7, bold=False,
            color=T.muted_rgb, align=PP_ALIGN.LEFT, rtl=False, vcenter=True)
    # خلفية المحتوى
    gradient_rect(slide, 0, HEADER_H, W, H-HEADER_H, T.bg, T.bg2, angle=90)
    # شريط سفلي — مرة واحدة فقط
    bt = rect(slide, 0, H-FOOTER_H, W, FOOTER_H, T.bg2_rgb)
    bta = rect(slide, 0, H-FOOTER_H, W, 0.06, T.accent_rgb)
    if bta: gradient_fill(bta, T.accent_grad1, T.accent_grad2, 0)


def _content_y():
    return HEADER_H + 0.28

def _content_h():
    return H - HEADER_H - FOOTER_H - 0.55


# ══════════════════════════════════════════════════════════════════════
# COVER — Classic
# ══════════════════════════════════════════════════════════════════════

def make_cover(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,angle=160)

    r_top=rect(slide,0,0,W,0.22,T.accent_rgb)
    if r_top: gradient_fill(r_top,T.accent_grad1,T.accent_grad2,0)
    r_bot=rect(slide,0,H-0.22,W,0.22,T.accent_rgb)
    if r_bot: gradient_fill(r_bot,T.accent_grad1,T.accent_grad2,0)
    vline(slide,0.22,0.22,H-0.44,T.accent_rgb,thickness=0.08)
    vline(slide,W-0.3,0.22,H-0.44,T.accent_rgb,thickness=0.08)

    if req.institution:
        txt(slide,req.institution,2.5,0.38,W-5.0,0.82,
            font=_FONT,size=12,bold=False,color=T.muted_rgb,
            align=PP_ALIGN.CENTER,rtl=True,vcenter=True)
    hline(slide,W*0.18,1.24,W*0.64,T.accent_rgb,thickness=0.05)

    # 42% عنوان + 55% جدول — توزيع محسّن
    title_y=1.42; total_h=H-0.22-title_y
    title_h=total_h*0.42; info_y=title_y+title_h+0.22
    info_h=H-0.22-info_y-0.12

    ts=26 if len(req.title_ar)<42 else 21 if len(req.title_ar)<65 else 17
    txt(slide,req.title_ar,2.2,title_y,W-4.4,title_h,
        font=_FONT,size=ts,bold=True,color=T.text_light_rgb,
        align=PP_ALIGN.CENTER,rtl=True,vcenter=True,line_spacing=1.2)
    if req.title_en:
        en_y=title_y+title_h+0.04
        txt(slide,req.title_en,2.5,en_y,W-5.0,0.72,
            font="Calibri",size=11,bold=False,italic=True,
            color=T.muted_rgb,align=PP_ALIGN.CENTER,rtl=False,vcenter=True)

    hl1=rect(slide,W*0.12,info_y-0.18,W*0.76,0.07,T.accent_rgb)
    if hl1: multi_stop_gradient(hl1,[(0,T.bg),(50,T.accent),(100,T.bg)],0)
    rect(slide,W*0.2,info_y-0.08,W*0.6,0.03,T.muted_rgb)

    rows=[("اسم الطالب",req.student_name)]
    if req.supervisor:     rows.append(("المشرف",req.supervisor))
    if req.co_supervisor:  rows.append(("المشرف المساعد",req.co_supervisor))
    if req.specialization: rows.append(("التخصص",req.specialization))
    if req.year:           rows.append(("السنة الجامعية",req.year))

    rh=info_h/max(len(rows),1)
    for i,(lbl,val) in enumerate(rows):
        y=info_y+i*rh
        fill=T.bg2_rgb if i%2==0 else T.card_rgb
        rb=rect(slide,MX,y,W-MX*2,rh-0.06,fill)
        acc=rect(slide,W-MX-0.18,y,0.18,rh-0.06,T.accent_rgb)
        if acc: set_solid_alpha(acc,70)
        txt(slide,lbl,MX+0.2,y,4.2,rh-0.06,
            font=_FONT,size=max(11,min(13,rh*7.5)),bold=True,
            color=T.accent_rgb,align=PP_ALIGN.RIGHT,rtl=True,vcenter=True,line_spacing=1.0)
        vline(slide,MX+4.5,y+rh*0.1,rh*0.7,T.muted_rgb,thickness=0.04)
        txt(slide,val,MX+4.7,y,W-MX*2-5.0,rh-0.06,
            font=_FONT,size=max(12,min(14.5,rh*9)),bold=False,
            color=T.text_light_rgb,align=PP_ALIGN.RIGHT,rtl=True,vcenter=True,line_spacing=1.0)

    return slide


def make_intro(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "مقدمة البحث", 1, req=req)

    cx = MX
    cw = W - MX * 2
    cy = _content_y()

    items = []
    if req.intro_overview:  items.append(("نظرة عامة", req.intro_overview))
    if req.intro_approach:  items.append(("المنهج المتبع", req.intro_approach))

    avail_h = _content_h()
    card_h = avail_h / max(len(items), 1) - 0.25

    for i, (lbl, val) in enumerate(items[:2]):
        y = cy + i * (card_h + 0.25)
        # بطاقة خلفية
        cb = rrect(slide, cx, y, cw - 0.4, card_h, T.bg2_rgb if i%2==0 else T.card_rgb, radius_pct=6)
        if cb:
            stops = [(0,T.bg2),(100,T.card)] if i%2==0 else [(0,T.card),(100,T.bg2)]
            multi_stop_gradient(cb, stops, 0)
        # خط accent يميني
        vline(slide, W - MX - 0.1, y, card_h, T.accent_rgb, thickness=0.12)
        # شريط عنوان
        hdr_bar = rrect(slide, cx, y, cw - 0.4, 0.62, T.accent_rgb, radius_pct=0)
        if hdr_bar: multi_stop_gradient(hdr_bar, [(0,T.accent),(100,T.accent2)], 0)
        # تسمية
        txt(slide, lbl, cx + 0.2, y, cw - 0.8, 0.62,
            font=_FONT, size=14, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)
        # خط تحت التسمية
        hline(slide, cx, y + 0.65, cw - 0.4, T.accent_rgb, thickness=0.04)
        # المحتوى
        txt(slide, val, cx + 0.2, y + 0.72, cw - 0.8, card_h - 0.82,
            font=_FONT, size=max(11, min(13, card_h * 4)), bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True, line_spacing=1.3)

    return slide


def make_plan(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "خطة البحث", page_num=2, req=req)

    CY = _content_y()
    CH = _content_h()
    chapters = req.chapters[:8]
    n = len(chapters)
    if not chapters: return slide

    gap = 0.14
    row_h = (CH - gap*(n-1)) / n  # يملأ الارتفاع كاملاً

    for i, ch in enumerate(chapters):
        y = CY + i*(row_h+gap)
        fill = T.bg2_rgb if i%2==0 else T.card_rgb
        rw = rrect(slide, MX, y, W-MX*2, row_h, fill, radius_pct=6)
        if rw:
            stops = [(0,T.bg2),(100,T.card)] if i%2==0 else [(0,T.card),(100,T.bg2)]
            multi_stop_gradient(rw, stops, 0)

        # خط acc يميني
        acc = rect(slide, W-MX-0.22, y, 0.22, row_h, T.accent_rgb)
        if acc: gradient_fill(acc, T.accent_grad1, T.accent_grad2, 90)

        # رقم الفصل
        num_label = f"الفصل {i+1}"
        txt(slide, num_label, MX+0.15, y, 3.2, row_h,
            font=_FONT, size=max(10, min(13, int(row_h*8))),
            bold=True, color=T.accent_rgb,
            align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)

        # فاصل عمودي
        vline(slide, MX+3.4, y+row_h*0.12, row_h*0.76, T.muted_rgb, thickness=0.04)

        # عنوان الفصل
        txt(slide, ch.title, MX+3.6, y, W-MX*2-4.5, row_h,
            font=_FONT, size=max(11, min(14, int(row_h*8.5))),
            bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)

        # الصفحات
        if ch.pages:
            txt(slide, ch.pages, MX+0.15, y, 1.8, row_h,
                font="Calibri", size=9, bold=False,
                color=T.muted_rgb, align=PP_ALIGN.LEFT, rtl=False)

    pass  # رقم الشريحة في الهيدر
    return slide

def make_problem(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "إشكالية البحث", 3, req=req)

    cx = MX
    cw = W - MX * 2
    cy = _content_y()

    if req.main_problem:
        # بطاقة الإشكالية
        vline(slide, W - MX - 0.14, cy, 2.6, T.accent_rgb, thickness=0.14)
        txt(slide, "الإشكالية الرئيسية", cx, cy, cw - 0.3, 0.65,
            font=_FONT, size=13, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.RIGHT, rtl=True)
        hline(slide, cx, cy + 0.68, cw - 0.3, T.muted_rgb, thickness=0.03)
        txt(slide, req.main_problem, cx, cy + 0.78, cw - 0.3, 1.7,
            font=_FONT, size=12, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)
        cy += 2.75

    if req.main_question:
        hline(slide, cx, cy, cw, T.accent_rgb, thickness=0.06)
        cy += 0.15
        txt(slide, "التساؤل الرئيسي", cx, cy, cw, 0.6,
            font=_FONT, size=12, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.RIGHT, rtl=True)
        txt(slide, req.main_question, cx, cy + 0.65, cw, 1.3,
            font=_FONT, size=12, bold=False, italic=True,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)
        cy += 2.1

    if req.sub_questions:
        hline(slide, cx, cy, cw, T.muted_rgb, thickness=0.03)
        cy += 0.2
        txt(slide, "التساؤلات الفرعية", cx, cy, cw, 0.55,
            font=_FONT, size=11, bold=True,
            color=T.muted_rgb, align=PP_ALIGN.RIGHT, rtl=True)
        cy += 0.6
        avail = H - cy - FOOTER_H - 0.25
        sub_h = min(avail / max(len(req.sub_questions), 1), 0.85)
        for i, q in enumerate(req.sub_questions[:6]):
            y = cy + i * sub_h
            txt(slide, f"{'─'} {q}", cx, y, cw, sub_h,
                font=_FONT, size=11, bold=False,
                color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    return slide


def make_objectives(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "أهداف البحث وفرضياته", page_num=4, req=req)
    CY = _content_y(); CH = _content_h()
    cols_data = []
    if req.objectives: cols_data.append(("الأهداف", req.objectives))
    if req.hypotheses:  cols_data.append(("الفرضيات", req.hypotheses))
    if not cols_data: return slide

    n_c = len(cols_data); gap = 0.3
    col_w = (W-MX*2-gap*(n_c-1)) / n_c

    for i, (lbl, items) in enumerate(cols_data[:2]):
        x = MX + i*(col_w+gap)
        # هيدر العمود
        hh = 0.65
        hdr = rect(slide, x, CY, col_w, hh, T.bg2_rgb)
        acc_top = rect(slide, x, CY, col_w, 0.1, T.accent_rgb)
        if acc_top: gradient_fill(acc_top, T.accent_grad1, T.accent_grad2, 0)
        acc_r = rect(slide, x+col_w-0.12, CY, 0.12, hh, T.accent_rgb)
        txt(slide, lbl, x+0.18, CY, col_w-0.4, hh,
            font=_FONT, size=14, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.RIGHT, rtl=True)

        # عناصر — تملأ الارتفاع
        ia = CH - hh - 0.1
        n_items = min(len(items), 8)
        ig = 0.1
        ih = (ia - ig*(n_items-1)) / n_items

        for j, item in enumerate(items[:8]):
            iy = CY + hh + 0.05 + j*(ih+ig)
            fill = T.bg2_rgb if j%2==0 else T.card_rgb
            rb = rect(slide, x, iy, col_w, ih, fill)

            # رقم + فاصل
            txt(slide, str(j+1), x+0.12, iy, 0.8, ih,
                font="Calibri", size=max(9, int(ih*6.5)),
                bold=True, color=T.accent_rgb,
                align=PP_ALIGN.CENTER, rtl=False)
            vline(slide, x+0.95, iy+ih*0.1, ih*0.8, T.muted_rgb, thickness=0.04)

            txt(slide, item, x+1.05, iy+0.04, col_w-1.25, ih-0.08,
                font=_FONT, size=max(9, min(11.5, ih*7)),
                bold=False, color=T.text_light_rgb,
                align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)

    pass  # رقم الشريحة في الهيدر
    return slide

def make_importance(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "أهمية البحث", page_num=5, req=req)
    CY = _content_y(); CH = _content_h()
    items = req.importance[:6]
    if not items: return slide
    n = len(items); gap = 0.14
    row_h = (CH - gap*(n-1)) / n

    for i, item in enumerate(items):
        y = CY + i*(row_h+gap)
        fill = T.bg2_rgb if i%2==0 else T.card_rgb
        rb = rect(slide, MX, y, W-MX*2, row_h, fill)
        acc = rect(slide, W-MX-0.2, y, 0.2, row_h, T.accent_rgb)
        if acc: gradient_fill(acc, T.accent_grad1, T.accent_grad2, 90)

        txt(slide, f"{i+1:02d}", MX+0.15, y, 1.4, row_h,
            font="Calibri", size=max(14, int(row_h*9)),
            bold=True, color=T.accent_rgb,
            align=PP_ALIGN.CENTER, rtl=False, vcenter=True)
        vline(slide, MX+1.6, y+row_h*0.08, row_h*0.84, T.muted_rgb, thickness=0.04)

        txt(slide, item, MX+1.8, y+0.08, W-MX*2-2.2, row_h-0.16,
            font=_FONT, size=max(10, min(12.5, row_h*7.5)),
            bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)

    pass  # رقم الشريحة في الهيدر
    return slide

def make_methodology(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "منهجية البحث", 6, req=req)

    cx = MX
    cw = W - MX * 2
    cy = _content_y()

    fields = []
    if req.methodology:  fields.append(("المنهج المتبع", req.methodology))
    if req.sample_type:  fields.append(("نوع العينة", req.sample_type))
    if req.sample_size:  fields.append(("حجم العينة", req.sample_size))
    if req.tool:         fields.append(("أداة الدراسة", req.tool))

    avail_h = _content_h()
    row_h = min(avail_h / max(len(fields), 1) - 0.15, 2.5)

    for i, (lbl, val) in enumerate(fields[:4]):
        y = cy + i * (row_h + 0.15)

        # الصف
        fill = T.bg2_rgb if i % 2 == 0 else T.card_rgb
        rect(slide, cx, y, cw, row_h, fill)

        # خط accent يميني
        vline(slide, W - MX - 0.12, y, row_h, T.accent_rgb, thickness=0.12)

        # التسمية
        rect(slide, cx, y, 5.0, row_h, T.bg2_rgb if i % 2 != 0 else T.card_rgb)
        txt(slide, lbl, cx + 0.15, y, 4.7, row_h,
            font=_FONT, size=12, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)

        # فاصل
        vline(slide, cx + 5.1, y + 0.1, row_h - 0.2, T.muted_rgb, thickness=0.04)

        # القيمة
        txt(slide, val, cx + 5.3, y + 0.1, cw - 5.7, row_h - 0.2,
            font=_FONT, size=12, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)

    return slide


def make_stats(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "الأرقام والإحصاءات الرئيسية", page_num=7, req=req)

    CY = _content_y()
    CH = _content_h()
    stats = req.stats[:6]
    if not stats: return slide

    cols = 3 if len(stats)>=3 else len(stats)
    rows_n = (len(stats)+cols-1)//cols
    gh = 0.28; gv = 0.2
    cw = (W-MX*2-gh*(cols-1)) / cols
    ch = (CH - gv*(rows_n-1)) / rows_n

    for i, stat in enumerate(stats):
        ci = i%cols; ri = i//cols
        x = MX+ci*(cw+gh); y = CY+ri*(ch+gv)

        cb = rect(slide, x, y, cw, ch, T.bg2_rgb)
        # خط acc أعلى
        tp = rect(slide, x, y, cw, 0.14, T.accent_rgb)
        if tp: gradient_fill(tp, T.accent_grad1, T.accent_grad2, 0)
        # خط acc يميني
        vline(slide, x+cw-0.12, y, ch, T.accent_rgb, thickness=0.12)

        # القيمة الرئيسية
        vs = 36 if len(stat.value)<=3 else 26 if len(stat.value)<=6 else 20
        txt(slide, stat.value, x+0.15, y+0.2, cw-0.4, ch*0.52,
            font="Calibri", size=vs, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)

        # فاصل
        hline(slide, x+cw*0.14, y+ch*0.58, cw*0.72, T.muted_rgb, thickness=0.04)

        if stat.unit:
            txt(slide, stat.unit, x+0.15, y+ch*0.6, cw-0.4, 0.48,
                font=_FONT, size=9.5, bold=False,
                color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

        txt(slide, stat.label, x+0.15, y+ch*0.73, cw-0.4, ch*0.25,
            font=_FONT, size=max(9, min(11, ch*5.5)),
            bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

    pass  # رقم الشريحة في الهيدر
    return slide

def make_results(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "نتائج البحث", page_num=8, req=req)
    CY = _content_y(); CH = _content_h()
    results = req.main_results[:8]
    n = len(results)
    if not results: return slide
    gap = 0.14
    row_h = (CH - gap*(n-1)) / n

    for i, result in enumerate(results):
        y = CY + i*(row_h+gap)
        fill = T.bg2_rgb if i%2==0 else T.card_rgb
        rb = rect(slide, MX, y, W-MX*2, row_h, fill)
        acc = rect(slide, W-MX-0.2, y, 0.2, row_h, T.accent_rgb)
        if acc:
            gradient_fill(acc, T.accent_grad1, T.accent_grad2, 90)
            set_solid_alpha(acc, max(18,56-i*7))

        txt(slide, str(i+1), MX+0.15, y, 0.9, row_h,
            font="Calibri", size=max(12, int(row_h*8)),
            bold=True, color=T.accent_rgb,
            align=PP_ALIGN.CENTER, rtl=False, vcenter=True)
        vline(slide, MX+1.1, y+row_h*0.08, row_h*0.84, T.muted_rgb, thickness=0.04)

        txt(slide, result, MX+1.3, y+0.07, W-MX*2-1.7, row_h-0.14,
            font=_FONT, size=max(10, min(12.5, row_h*7.5)),
            bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)

    pass  # رقم الشريحة في الهيدر
    return slide

def make_conclusion(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "خاتمة البحث", page_num=9, req=req)
    CY = _content_y(); CH = _content_h()
    cw = W - MX*2

    cb = rect(slide, MX, CY, cw, CH, T.bg2_rgb)
    acc_r = rect(slide, W-MX-0.2, CY, 0.2, CH, T.accent_rgb)
    if acc_r: gradient_fill(acc_r, T.accent_grad1, T.accent_grad2, 90)
    acc_l = rect(slide, MX, CY, 0.12, CH, T.bg2_rgb)

    txt(slide, "الاستنتاج العام", MX+0.3, CY+0.18, cw-0.6, 0.72,
        font=_FONT, size=14, bold=True,
        color=T.accent_rgb, align=PP_ALIGN.RIGHT, rtl=True)

    hl = rect(slide, MX, CY+0.95, cw, 0.07, T.accent_rgb)
    if hl: gradient_fill(hl, T.accent_grad1, T.accent_grad2, 0)

    txt(slide, req.general_conclusion, MX+0.3, CY+1.1, cw-0.6, CH-2.0,
        font=_FONT, size=max(11,min(14,int((CH-2.0)*5))), bold=False,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True, line_spacing=1.35)

    ny = CY + CH - 0.78
    hline(slide, MX+cw*0.18, ny, cw*0.64, T.accent_rgb, thickness=0.06)
    txt(slide, req.student_name, MX, ny+0.1, cw, 0.62,
        font=_FONT, size=13, bold=True,
        color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=True)

    pass  # رقم الشريحة في الهيدر
    return slide

def make_recommendations(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "توصيات البحث", page_num=10, req=req)
    CY = _content_y(); CH = _content_h()
    recs = req.recommendations[:8]
    n = len(recs)
    if not recs: return slide
    gap = 0.14
    row_h = (CH - gap*(n-1)) / n

    for i, rec in enumerate(recs):
        y = CY + i*(row_h+gap)
        fill = T.bg2_rgb if i%2==0 else T.card_rgb
        rb = rect(slide, MX, y, W-MX*2, row_h, fill)
        acc = rect(slide, W-MX-0.2, y, 0.2, row_h, T.accent_rgb)
        if acc: gradient_fill(acc, T.accent_grad1, T.accent_grad2, 90)
        dot = oval(slide, MX+0.25, y+(row_h-0.38)/2, 0.38, 0.38, T.accent_rgb)
        if dot: gradient_fill(dot, T.accent_grad1, T.accent_grad2, 135)
        txt(slide, rec, MX+0.8, y+0.05, W-MX*2-1.2, row_h-0.1,
            font=_FONT, size=max(10, min(12.5, row_h*7.5)),
            bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)

    pass  # رقم الشريحة في الهيدر
    return slide

def make_future(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "آفاق البحث المستقبلية", page_num=11, req=req)
    CY = _content_y(); CH = _content_h()
    items = req.future_work[:6]
    if not items: return slide
    cols = 2 if len(items)>3 else 1
    rows_n = (len(items)+cols-1)//cols
    gh = 0.25; gv = 0.18
    cw = (W-MX*2-gh*(cols-1)) / cols
    ch = (CH - gv*(rows_n-1)) / rows_n

    for i, item in enumerate(items):
        ci = i%cols; ri = i//cols
        x = MX+ci*(cw+gh); y = CY+ri*(ch+gv)
        cb = rect(slide, x, y, cw, ch, T.bg2_rgb if i%2==0 else T.card_rgb)
        acc = rect(slide, x+cw-0.18, y, 0.18, ch, T.accent_rgb)
        if acc: gradient_fill(acc, T.accent_grad1, T.accent_grad2, 90)
        tp = rect(slide, x, y, cw, 0.1, T.accent_rgb)
        if tp: gradient_fill(tp, T.accent_grad1, T.accent_grad2, 0)

        txt(slide, str(i+1), x+0.15, y, 1.0, ch,
            font="Calibri", size=max(16, int(ch*8)),
            bold=True, color=T.accent_rgb,
            align=PP_ALIGN.CENTER, rtl=False, vcenter=True)
        vline(slide, x+1.15, y+ch*0.1, ch*0.8, T.muted_rgb, thickness=0.04)

        txt(slide, item, x+1.3, y+0.1, cw-1.65, ch-0.2,
            font=_FONT, size=max(10, min(12.5, ch*6)),
            bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)

    pass  # رقم الشريحة في الهيدر
    return slide

def make_references(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _header(slide, T, "قائمة المراجع", page_num=12, req=req)
    CY = _content_y(); CH = _content_h()
    refs = req.references[:12]
    n = len(refs)
    if not refs: return slide
    gap = 0.12
    row_h = (CH - gap*(n-1)) / n

    for i, ref in enumerate(refs):
        y = CY + i*(row_h+gap)
        fill = T.bg2_rgb if i%2==0 else T.card_rgb
        rb = rect(slide, MX, y, W-MX*2, row_h, fill)
        acc = rect(slide, W-MX-0.18, y, 0.18, row_h, T.accent_rgb)
        if acc: set_solid_alpha(acc, 60)

        txt(slide, f"[{i+1}]", MX+0.12, y, 0.72, row_h,
            font="Calibri", size=max(8, int(row_h*6.5)),
            bold=True, color=T.accent_rgb,
            align=PP_ALIGN.CENTER, rtl=False, vcenter=True)
        vline(slide, MX+0.9, y+row_h*0.08, row_h*0.84, T.muted_rgb, thickness=0.03)

        txt(slide, ref, MX+1.05, y+0.03, W-MX*2-1.4, row_h-0.06,
            font=_FONT, size=max(9, min(11, row_h*7)),
            bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)

    pass  # رقم الشريحة في الهيدر
    return slide

def make_final(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    gradient_rect(slide, 0, 0, W, H, T.grad1, T.grad2, angle=160)

    # إطار خارجي
    r_top = rect(slide, 0, 0, W, 0.22, T.accent_rgb)
    if r_top: gradient_fill(r_top, T.accent_grad1, T.accent_grad2, 0)
    r_bot = rect(slide, 0, H-0.22, W, 0.22, T.accent_rgb)
    if r_bot: gradient_fill(r_bot, T.accent_grad1, T.accent_grad2, 0)
    vline(slide, 0.22, 0.22, H-0.44, T.accent_rgb, thickness=0.08)
    vline(slide, W-0.3, 0.22, H-0.44, T.accent_rgb, thickness=0.08)

    cw = W - MX*2
    # بطاقة مركزية تشغل 75% من الشريحة
    cy = H*0.1; ch = H*0.8
    cb = rect(slide, MX, cy, cw, ch, T.bg2_rgb)
    if cb: multi_stop_gradient(cb,[(0,T.bg2),(100,T.card)],135)
    acc_top = rect(slide, MX, cy, cw, 0.16, T.accent_rgb)
    if acc_top: gradient_fill(acc_top, T.accent_grad1, T.accent_grad2, 0)
    acc_bot = rect(slide, MX, cy+ch-0.16, cw, 0.16, T.accent_rgb)
    if acc_bot: gradient_fill(acc_bot, T.accent_grad1, T.accent_grad2, 0)
    acc_r = rect(slide, MX+cw-0.18, cy, 0.18, ch, T.accent_rgb)
    acc_l = rect(slide, MX, cy, 0.18, ch, T.accent_rgb)

    txt(slide, "شكراً وتقديراً", MX+0.3, cy+0.35, cw-0.6, H*0.25,
        font=_FONT, size=34, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.CENTER, rtl=True)

    hl1 = rect(slide, MX+cw*0.15, cy+H*0.3, cw*0.7, 0.07, T.accent_rgb)
    if hl1: multi_stop_gradient(hl1,[(0,T.bg2),(50,T.accent),(100,T.bg2)],0)
    rect(slide, MX+cw*0.25, cy+H*0.3+0.16, cw*0.5, 0.03, T.muted_rgb)

    txt(slide, req.student_name, MX+0.3, cy+H*0.33, cw-0.6, H*0.14,
        font=_FONT, size=20, bold=True,
        color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=True)

    ts = req.title_ar[:70]+("..." if len(req.title_ar)>70 else "")
    txt(slide, ts, MX+0.5, cy+H*0.48, cw-1.0, H*0.18,
        font=_FONT, size=12, bold=False, italic=True,
        color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=True)

    footer = []
    if req.institution: footer.append(req.institution)
    if req.year: footer.append(req.year)
    if footer:
        txt(slide, "  ·  ".join(footer), MX+0.3, cy+H*0.68, cw-0.6, H*0.1,
            font=_FONT, size=11, bold=False,
            color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=True)

    return slide
