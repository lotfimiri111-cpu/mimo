"""
Premium Engine v17.3 — شريط جانبي + محتوى يملأ الشريحة 100%
"""
from __future__ import annotations
from pptx.enum.text import PP_ALIGN
from engine.primitives import (
    W, H, rect, rrect, oval, bg, hline, vline,
    gradient_fill, gradient_rect, shadow, set_solid_alpha,
    multi_stop_gradient, glow, diamond, decorative_dots,
    number_badge, icon_circle, slide_number, txt, blank_slide,
)
from core.themes import Theme
from core.models import PresentationRequest

_FONT = "Cairo"
def set_font(n): global _FONT; _FONT = n

SW = 5.2   # عرض الشريط الجانبي — ثابت لكل الشرائح
CX = SW + 0.55  # بداية منطقة المحتوى
CW = W - CX - 0.55  # عرض منطقة المحتوى

# ─── شريط جانبي ثابت لكل الشرائح ─────────────────────────────────────
def _section_slide(prs, T, icon, label1, label2=""):
    slide = blank_slide(prs)
    bg(slide, T.bg_rgb)
    _sidebar(slide, T, icon, label1, label2)
    return slide


def _sidebar(slide, T, icon, label1, label2=""):
    # خلفية الشريط بتدرج
    sb = gradient_rect(slide, 0, 0, SW, H, T.grad2, T.grad1, angle=180)
    # خط فاصل
    sep = rect(slide, SW, 0, 0.12, H, T.accent_rgb)
    if sep: gradient_fill(sep, T.accent_grad1, T.accent_grad2, 90)
    # زخارف الشريط
    oval(slide, -3, -3, 9, 9, T.accent_rgb, alpha=6)
    oval(slide, 0.5, H-5.5, 6.5, 6.5, T.bg2_rgb, alpha=45)
    decorative_dots(slide, 0.4, H*0.55, 5, 4, 0.14, 0.32, T.accent_rgb, alpha=12)

    # دائرة الأيقونة في المنتصف العلوي من الشريط
    ic_y = H*0.18
    ic_bg = oval(slide, SW/2-1.5, ic_y, 3.0, 3.0, T.accent_rgb)
    if ic_bg:
        multi_stop_gradient(ic_bg,[(0,T.accent),(100,T.accent2)],135)
        shadow(ic_bg, blur=14, dist=4, alpha=0.4)
        glow(ic_bg, T.accent.lstrip('#'), radius=20, alpha=0.2)
    txt(slide, icon, SW/2-1.5, ic_y+0.3, 3.0, 2.4,
        font="Calibri", size=34, bold=False,
        color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False)

    # عنوان القسم في الشريط
    label_y = ic_y + 3.2
    txt(slide, label1, 0.2, label_y, SW-0.4, 1.1,
        font=_FONT, size=16, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.CENTER, rtl=True)
    if label2:
        txt(slide, label2, 0.2, label_y+1.15, SW-0.4, 0.9,
            font=_FONT, size=12, bold=False,
            color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=True)

    # خط فاصل في الشريط
    hl = rect(slide, 0.5, label_y+2.2, SW-1.0, 0.05, T.accent_rgb)
    if hl: multi_stop_gradient(hl,[(0,T.bg),(50,T.accent),(100,T.bg)],0)

    # خلفية منطقة المحتوى
    rect(slide, SW+0.12, 0, W-SW-0.12, H, T.bg2_rgb)
    # تدرج خفيف على منطقة المحتوى
    gr = gradient_rect(slide, SW+0.12, 0, W-SW-0.12, H, T.grad1, T.bg2, angle=135)
    # زخارف المحتوى
    oval(slide, W-8, -3, 11, 11, T.accent_rgb, alpha=4)
    diamond(slide, W-5, H-5, 4, 4, T.accent_rgb, alpha=5)

# ─── COVER ─────────────────────────────────────────────────────────────

def make_cover(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    gradient_rect(slide,0,0,SW+0.5,H,T.grad2,T.grad1,angle=180)
    sepc=rect(slide,SW+0.5,0,0.12,H,T.accent_rgb)
    if sepc: gradient_fill(sepc,T.accent_grad1,T.accent_grad2,90)
    oval(slide,-3,-3,9,9,T.accent_rgb,alpha=6)
    oval(slide,0.5,H-5.5,7,7,T.bg2_rgb,alpha=45)

    ic_bg=oval(slide,SW+0.5-3.3,H*0.08,3.3,3.3,T.accent_rgb)
    if ic_bg:
        multi_stop_gradient(ic_bg,[(0,T.accent),(100,T.accent2)],135)
        shadow(ic_bg,blur=14,dist=4,alpha=0.42)
        glow(ic_bg,T.accent.lstrip("#"),radius=20,alpha=0.2)
    txt(slide,"🎓",SW+0.5-3.3,H*0.08+0.35,3.3,2.7,
        font="Calibri",size=34,bold=False,color=T.text_dark_rgb,
        align=PP_ALIGN.CENTER,rtl=False,vcenter=True)

    if req.institution:
        txt(slide,req.institution,0.2,H*0.46,SW+0.28,2.2,
            font=_FONT,size=10,bold=False,color=T.muted_rgb,
            align=PP_ALIGN.CENTER,rtl=True,vcenter=True,line_spacing=1.2)
    if req.year:
        yb=rrect(slide,0.4,H-1.4,SW-0.32,0.62,T.accent_rgb,radius_pct=40)
        if yb: multi_stop_gradient(yb,[(0,T.accent),(100,T.accent2)],0)
        txt(slide,req.year,0.4,H-1.4,SW-0.32,0.62,
            font="Calibri",size=13,bold=True,color=T.text_dark_rgb,
            align=PP_ALIGN.CENTER,rtl=False,vcenter=True)

    gradient_rect(slide,SW+0.62,0,W-SW-0.62,H,T.grad1,T.bg2,angle=135)
    oval(slide,W-8,-3,11,11,T.accent_rgb,alpha=4)
    diamond(slide,W-5,H-5,4,4,T.accent_rgb,alpha=5)

    mcx=SW+1.0; mcw=W-mcx-0.8
    tp=rect(slide,mcx-0.3,0,mcw+0.3,0.38,T.accent_rgb)
    if tp: multi_stop_gradient(tp,[(0,T.bg),(40,T.accent),(60,T.accent2),(100,T.bg)],0)

    # عنوان بارتفاع ثابت + معلومات تملأ الباقي
    title_y=0.55; title_h=7.2
    info_y=title_y+title_h+0.25; info_h=H-info_y-0.36

    tc=rrect(slide,mcx,title_y,mcw,title_h,T.card_rgb,radius_pct=12)
    if tc:
        multi_stop_gradient(tc,[(0,T.card),(100,T.bg2)],135)
        shadow(tc,blur=22,dist=6,alpha=0.48)
    ct=rrect(slide,mcx,title_y,mcw,0.32,T.accent_rgb,radius_pct=0)
    if ct: multi_stop_gradient(ct,[(0,T.accent),(50,T.accent2),(100,T.accent)],0)
    vline(slide,mcx+mcw-0.2,title_y+0.32,title_h-0.32,T.accent_rgb,thickness=0.2)

    ts=24 if len(req.title_ar)<40 else 19 if len(req.title_ar)<65 else 16
    txt(slide,req.title_ar,mcx+0.4,title_y+0.38,mcw-0.85,title_h*0.62,
        font=_FONT,size=ts,bold=True,color=T.text_light_rgb,
        align=PP_ALIGN.CENTER,rtl=True,vcenter=True,line_spacing=1.2)
    if req.title_en:
        txt(slide,req.title_en,mcx+0.4,title_y+title_h*0.64,mcw-0.85,title_h*0.2,
            font="Calibri",size=10.5,bold=False,italic=True,
            color=T.muted_rgb,align=PP_ALIGN.CENTER,rtl=False,vcenter=True)
    hl=rect(slide,mcx+mcw*0.08,title_y+title_h*0.84,mcw*0.84,0.05,T.accent_rgb)
    if hl: multi_stop_gradient(hl,[(0,T.bg2),(50,T.accent),(100,T.bg2)],0)

    ic=rrect(slide,mcx,info_y,mcw,info_h,T.card_rgb,radius_pct=12)
    if ic:
        multi_stop_gradient(ic,[(0,T.bg2),(100,T.card)],135)
        shadow(ic,blur=14,dist=4,alpha=0.34)
    vline(slide,mcx+mcw-0.18,info_y,info_h,T.accent_rgb,thickness=0.18)

    rows=[("الطالب",req.student_name)]
    if req.supervisor:     rows.append(("المشرف",req.supervisor))
    if req.co_supervisor:  rows.append(("المشرف المساعد",req.co_supervisor))
    if req.specialization: rows.append(("التخصص",req.specialization))
    if req.year:           rows.append(("السنة",req.year))

    rh=info_h/max(len(rows),1)
    for i,(lbl,val) in enumerate(rows):
        y=info_y+i*rh
        rb=rrect(slide,mcx+0.25,y+0.04,mcw-0.62,rh-0.08,T.bg_rgb,radius_pct=7)
        if rb: set_solid_alpha(rb,50)
        txt(slide,f"{lbl} :",mcx+0.42,y+0.04,4.5,rh-0.08,
            font=_FONT,size=max(10.5,min(12.5,rh*7.5)),bold=True,
            color=T.accent_rgb,align=PP_ALIGN.RIGHT,rtl=True,vcenter=True)
        vline(slide,mcx+5.15,y+rh*0.12,rh*0.76,T.muted_rgb,thickness=0.04)
        txt(slide,val,mcx+5.35,y+0.04,mcw-6.0,rh-0.08,
            font=_FONT,size=max(12,min(14.5,rh*9)),bold=False,
            color=T.text_light_rgb,align=PP_ALIGN.RIGHT,rtl=True,vcenter=True)

    fb=rrect(slide,mcx,H-0.33,mcw,0.28,T.bg_rgb,radius_pct=0)
    if fb: set_solid_alpha(fb,45)
    txt(slide,"✦  مذكرتي Pro",mcx+0.3,H-0.33,mcw-0.6,0.28,
        font=_FONT,size=11,bold=False,italic=True,color=T.muted_rgb,
        align=PP_ALIGN.RIGHT,rtl=True,vcenter=True)

    bt=rect(slide,0,H-0.28,W,0.28,T.accent_rgb)
    if bt: gradient_fill(bt,T.accent_grad1,T.accent_grad2,0)
    return slide


# ─── PLAN ──────────────────────────────────────────────────────────────
def make_plan(prs, req: PresentationRequest, T: Theme):
    slide = _section_slide(prs, T, "📋", "خطة", "البحث")
    chapters = req.chapters[:8]
    n = len(chapters)
    if not chapters: return slide

    # يملأ الشريحة من أعلى إلى أسفل
    CY = 0.28; CH = H - CY - 0.22
    gap = 0.14
    row_h = (CH - gap*(n-1)) / n

    for i, ch in enumerate(chapters):
        y = CY + i*(row_h+gap)
        even = i%2==0

        # خلفية الصف بالكامل
        rw = rect(slide, CX, y, CW, row_h,
                  T.card_rgb if even else T.bg2_rgb)
        if rw:
            stops = [(0,T.card),(100,T.bg2)] if even else [(0,T.bg2),(100,T.card)]
            multi_stop_gradient(rw, stops, 0)

        # شريط acc يميني
        acc = rect(slide, CX+CW-0.22, y, 0.22, row_h, T.accent_rgb)
        if acc: gradient_fill(acc, T.accent_grad1, T.accent_grad2, 90)

        # دائرة رقم
        nd = min(0.68, row_h*0.72)
        nx = CX+CW-1.7; ny = y+(row_h-nd)/2
        nc = oval(slide, nx, ny, nd, nd, T.accent_rgb)
        if nc:
            multi_stop_gradient(nc,[(0,T.accent),(100,T.accent2)],135)
            shadow(nc, blur=6, dist=2, alpha=0.28)
        txt(slide, str(i+1), nx, ny, nd, nd,
            font="Calibri", size=max(9, int(nd*10)),
            bold=True, color=T.text_dark_rgb,
            align=PP_ALIGN.CENTER, rtl=False)

        # عنوان الفصل
        txt(slide, ch.title, CX+0.25, y, CW-2.45, row_h,
            font=_FONT, size=max(11, min(15, int(row_h*9))),
            bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)

        # الصفحات
        if ch.pages:
            pb = rrect(slide, CX+0.22, y+(row_h-0.34)/2, 1.85, 0.34,
                       T.bg_rgb, radius_pct=40)
            if pb: set_solid_alpha(pb, 52)
            txt(slide, ch.pages, CX+0.22, y+(row_h-0.34)/2, 1.85, 0.34,
                font="Calibri", size=8.5, bold=False,
                color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=False)

    slide_number(slide, 2, getattr(req, "_total_slides", 13), T)
    return slide


# ─── PROBLEM ───────────────────────────────────────────────────────────
def make_problem(prs, req: PresentationRequest, T: Theme):
    slide = _section_slide(prs, T, "❓", "إشكالية", "البحث")
    CY = 0.35; CH = H - CY - 0.3

    sections = []
    if req.main_problem: sections.append('p')
    if req.main_question: sections.append('q')
    if req.sub_questions: sections.append('s')

    weights = {'p':2.8,'q':1.6,'s':2.0}
    total_w = sum(weights[s] for s in sections)
    gap = 0.22
    avail = CH - gap*(len(sections)-1)

    cy = CY
    if 'p' in sections:
        h = avail * weights['p']/total_w
        pc = rrect(slide, CX, cy, CW, h, T.card_rgb, radius_pct=12)
        if pc:
            multi_stop_gradient(pc,[(0,T.card),(100,T.bg2)],135)
            shadow(pc, blur=16, dist=5, alpha=0.4)
            glow(pc, T.accent.lstrip('#'), radius=20, alpha=0.09)
        lb = rrect(slide, CX+CW-6.2, cy, 5.8, 0.5, T.accent_rgb, radius_pct=0)
        if lb: multi_stop_gradient(lb,[(0,T.accent),(100,T.accent2)],0)
        txt(slide, "◆  الإشكالية الرئيسية", CX+CW-6.2, cy, 5.8, 0.5,
            font=_FONT, size=11, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=True)
        txt(slide, "❝", CX+0.3, cy+0.58, 1.3, 1.0,
            font="Calibri", size=28, bold=False,
            color=T.accent_rgb, align=PP_ALIGN.LEFT, rtl=False)
        txt(slide, req.main_problem, CX+1.75, cy+0.55, CW-2.2, h-0.72,
            font=_FONT, size=12.5, bold=False,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)
        cy += h+gap

    if 'q' in sections:
        h = avail * weights['q']/total_w
        qc = rrect(slide, CX, cy, CW, h, T.bg2_rgb, radius_pct=10)
        if qc: shadow(qc, blur=8, dist=2, alpha=0.25)
        vline(slide, CX+CW-0.2, cy, h, T.accent_rgb, thickness=0.2)
        dot = oval(slide, CX+CW-2.8, cy+h/2-0.18, 0.38, 0.38, T.accent_rgb)
        if dot: multi_stop_gradient(dot,[(0,T.accent),(100,T.accent2)],135)
        txt(slide, req.main_question, CX+0.3, cy, CW-1.5, h,
            font=_FONT, size=12.5, bold=True, italic=True,
            color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)
        cy += h+gap

    if 's' in sections and req.sub_questions:
        h = avail * weights['s']/total_w
        subs = req.sub_questions[:5]
        sub_h = h / len(subs)
        for i, q in enumerate(subs):
            y2 = cy+i*sub_h
            if i%2==0:
                sb = rrect(slide, CX, y2, CW, sub_h-0.05, T.card_rgb, radius_pct=5)
                if sb: set_solid_alpha(sb, 48)
            nc = oval(slide, CX+CW-2.5, y2+(sub_h-0.33)/2, 0.33, 0.33, T.accent_rgb)
            if nc: set_solid_alpha(nc, 65)
            txt(slide, str(i+1), CX+CW-2.5, y2+(sub_h-0.33)/2, 0.33, 0.33,
                font="Calibri", size=7.5, bold=True,
                color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)
            txt(slide, q, CX+0.3, y2, CW-3.1, sub_h,
                font=_FONT, size=max(10, min(12, sub_h*8)),
                bold=False, color=T.muted_rgb,
                align=PP_ALIGN.RIGHT, rtl=True)

    slide_number(slide, 3, getattr(req, "_total_slides", 13), T)
    return slide

# ─── OBJECTIVES ────────────────────────────────────────────────────────
def make_objectives(prs, req: PresentationRequest, T: Theme):
    slide = _section_slide(prs, T, "🎯", "أهداف", "وفرضيات")
    CY = 0.35; CH = H - CY - 0.3
    cols = []
    if req.objectives: cols.append(("🎯  الأهداف", req.objectives))
    if req.hypotheses:  cols.append(("💡  الفرضيات", req.hypotheses))
    if not cols: return slide

    n_c = len(cols)
    gap = 0.25
    col_w = (CW - gap*(n_c-1)) / n_c

    for i, (lbl, items) in enumerate(cols):
        x = CX + i*(col_w+gap)
        cc = rrect(slide, x, CY, col_w, CH, T.card_rgb, radius_pct=12)
        if cc:
            multi_stop_gradient(cc,[(0,T.card),(100,T.bg2)],150)
            shadow(cc, blur=14, dist=4, alpha=0.35)
        hh = 0.68
        hdr = rrect(slide, x, CY, col_w, hh, T.accent_rgb, radius_pct=0)
        if hdr: multi_stop_gradient(hdr,[(0,T.accent2),(100,T.accent)],0)
        txt(slide, lbl, x+0.18, CY, col_w-0.36, hh,
            font=_FONT, size=14, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=True)

        ia = CH - hh - 0.12
        n_items = min(len(items), 8)
        ig = 0.1
        ih = (ia - ig*(n_items-1)) / n_items

        for j, item in enumerate(items[:8]):
            iy = CY + hh + 0.06 + j*(ih+ig)
            rb = rrect(slide, x+0.1, iy, col_w-0.2, ih,
                       T.bg2_rgb if j%2==0 else T.bg_rgb, radius_pct=7)
            if rb: set_solid_alpha(rb, 72)
            number_badge(slide, x+col_w-0.78, iy+(ih-0.5)/2, 0.5, j+1, T)
            txt(slide, item, x+0.22, iy+0.05, col_w-1.2, ih-0.1,
                font=_FONT, size=max(9, min(11.5, ih*7.5)),
                bold=False, color=T.text_light_rgb,
                align=PP_ALIGN.RIGHT, rtl=True)

    slide_number(slide, 4, getattr(req, "_total_slides", 13), T)
    return slide

# ─── IMPORTANCE ────────────────────────────────────────────────────────
def make_importance(prs, req: PresentationRequest, T: Theme):
    slide = _section_slide(prs, T, "⭐", "أهمية", "البحث")
    CY = 0.35; CH = H - CY - 0.3
    items = (req.importance or [])[:6]
    if not items: return slide

    icons = ["⭐","🔑","📌","🌟","💎","🏆"]
    cols = 2 if len(items)>3 else 1
    rows_n = (len(items)+cols-1)//cols
    gh = 0.18; gv = 0.18
    cw = (CW - gh*(cols-1)) / cols
    ch = (CH - gv*(rows_n-1)) / rows_n

    for i, item in enumerate(items):
        ci = i%cols; ri = i//cols
        x = CX+ci*(cw+gh); y = CY+ri*(ch+gv)
        cc = rrect(slide, x, y, cw, ch, T.card_rgb, radius_pct=10)
        if cc:
            multi_stop_gradient(cc,[(0,T.card),(100,T.bg2)],145)
            shadow(cc, blur=12, dist=3, alpha=0.32)
        acc = rect(slide, x+cw-0.25, y, 0.25, ch, T.accent_rgb)
        if acc: gradient_fill(acc, T.accent_grad1, T.accent_grad2, 90)

        ic_s = min(1.3, ch*0.6)
        icon_circle(slide, x+0.25, y+(ch-ic_s)/2, ic_s,
                    T.accent_grad1, T.accent_grad2,
                    icons[i%len(icons)], max(13, int(ic_s*11)), T)
        txt(slide, item, x+ic_s+0.5, y+0.1, cw-ic_s-1.0, ch-0.2,
            font=_FONT, size=max(9.5, min(12.5, ch*7)),
            bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True, vcenter=True)

    slide_number(slide, 5, getattr(req, "_total_slides", 13), T)
    return slide

# ─── METHODOLOGY ───────────────────────────────────────────────────────
def make_methodology(prs, req: PresentationRequest, T: Theme):
    slide = _section_slide(prs, T, "🔬", "منهجية", "البحث")
    CY = 0.35; CH = H - CY - 0.3
    icons_map = {"المنهج":"📊","العينة":"👥","حجم العينة":"📏","الأداة":"🛠️"}
    fields = []
    if req.methodology: fields.append(("المنهج", req.methodology))
    if req.sample_type:  fields.append(("العينة", req.sample_type))
    if req.sample_size:  fields.append(("حجم العينة", req.sample_size))
    if req.tool:         fields.append(("الأداة", req.tool))
    if not fields: return slide

    cols = 2 if len(fields)>2 else len(fields)
    rows_n = (len(fields)+cols-1)//cols
    gh = 0.22; gv = 0.2
    cw = (CW - gh*(cols-1)) / cols
    ch = (CH - gv*(rows_n-1)) / rows_n

    for i, (lbl, val) in enumerate(fields[:4]):
        ci = i%cols; ri = i//cols
        x = CX+ci*(cw+gh); y = CY+ri*(ch+gv)
        cc = rrect(slide, x, y, cw, ch, T.card_rgb, radius_pct=12)
        if cc:
            multi_stop_gradient(cc,[(0,T.card),(100,T.bg2)],145)
            shadow(cc, blur=13, dist=4, alpha=0.38)

        ic_s = min(1.8, ch*0.36)
        ic_x = x+cw/2-ic_s/2
        ic_bg = oval(slide, ic_x, y+0.28, ic_s, ic_s, T.accent_rgb)
        if ic_bg:
            multi_stop_gradient(ic_bg,[(0,T.accent),(100,T.accent2)],135)
            shadow(ic_bg, blur=9, dist=3, alpha=0.3)
            glow(ic_bg, T.accent.lstrip('#'), radius=14, alpha=0.16)
        txt(slide, icons_map.get(lbl,"📌"), ic_x, y+0.32, ic_s, ic_s*0.86,
            font="Calibri", size=max(14, int(ic_s*10)),
            bold=False, color=T.text_dark_rgb,
            align=PP_ALIGN.CENTER, rtl=False)

        lbl_y = y + ic_s + 0.44
        txt(slide, lbl, x+0.2, lbl_y, cw-0.4, 0.62,
            font=_FONT, size=12.5, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=True)
        hl = rect(slide, x+cw*0.14, lbl_y+0.65, cw*0.72, 0.04, T.muted_rgb)
        txt(slide, val, x+0.2, lbl_y+0.76, cw-0.4, ch-(lbl_y-y)-0.92,
            font=_FONT, size=max(9, min(11, ch*5)),
            bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.CENTER, rtl=True, vcenter=True)

    slide_number(slide, 6, getattr(req, "_total_slides", 13), T)
    return slide

# ─── STATS ─────────────────────────────────────────────────────────────
def make_stats(prs, req: PresentationRequest, T: Theme):
    slide = _section_slide(prs, T, "📈", "الأرقام", "الرئيسية")
    CY = 0.35; CH = H - CY - 0.3
    stats = req.stats[:6]
    if not stats: return slide

    cols = 3 if len(stats)>=3 else len(stats)
    rows_n = (len(stats)+cols-1)//cols
    gh = 0.22; gv = 0.2
    cw = (CW - gh*(cols-1)) / cols
    ch = (CH - gv*(rows_n-1)) / rows_n

    for i, stat in enumerate(stats):
        ci = i%cols; ri = i//cols
        x = CX+ci*(cw+gh); y = CY+ri*(ch+gv)

        cc = rrect(slide, x, y, cw, ch, T.card_rgb, radius_pct=12)
        if cc:
            multi_stop_gradient(cc,[(0,T.bg2),(50,T.card),(100,T.bg2)],135)
            shadow(cc, blur=16, dist=4, alpha=0.4)

        tp = rrect(slide, x, y, cw, 0.28, T.accent_rgb, radius_pct=0)
        if tp:
            multi_stop_gradient(tp,[(0,T.accent2),(50,T.accent),(100,T.accent2)],0)
            glow(tp, T.accent.lstrip('#'), radius=9, alpha=0.25)

        bp = rrect(slide, x, y+ch-0.2, cw, 0.2, T.accent_rgb, radius_pct=0)
        if bp: set_solid_alpha(bp, 35)

        vs = 36 if len(stat.value)<=3 else 28 if len(stat.value)<=6 else 20
        txt(slide, stat.value, x+0.15, y+0.32, cw-0.3, ch*0.52,
            font="Calibri", size=vs, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)

        if stat.unit:
            ub = rrect(slide, x+cw/2-1.5, y+ch*0.53+0.1, 3.0, 0.42,
                       T.bg_rgb, radius_pct=40)
            if ub: set_solid_alpha(ub, 52)
            txt(slide, stat.unit, x+cw/2-1.5, y+ch*0.53+0.1, 3.0, 0.42,
                font=_FONT, size=9.5, bold=False,
                color=T.muted_rgb, align=PP_ALIGN.CENTER, rtl=True)

        hl = rect(slide, x+cw*0.14, y+ch*0.7, cw*0.72, 0.04, T.muted_rgb)
        txt(slide, stat.label, x+0.15, y+ch*0.72, cw-0.3, ch*0.26,
            font=_FONT, size=max(9, min(11, ch*5.5)),
            bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.CENTER, rtl=True)

    slide_number(slide, 7, getattr(req, "_total_slides", 13), T)
    return slide

# ─── RESULTS ───────────────────────────────────────────────────────────
def make_results(prs, req: PresentationRequest, T: Theme):
    slide = _section_slide(prs, T, "📊", "نتائج", "البحث")
    CY = 0.35; CH = H - CY - 0.3
    results = req.main_results[:8]
    n = len(results)
    if not results: return slide

    gap = 0.14
    row_h = (CH - gap*(n-1)) / n

    for i, result in enumerate(results):
        y = CY + i*(row_h+gap)
        even = i%2==0
        rw = rrect(slide, CX, y, CW, row_h,
                   T.card_rgb if even else T.bg2_rgb, radius_pct=8)
        if rw:
            stops = [(0,T.card),(100,T.bg2)] if even else [(0,T.bg2),(100,T.card)]
            multi_stop_gradient(rw, stops, 0)
            shadow(rw, blur=4, dist=1, alpha=0.16)
        acc = rect(slide, CX+CW-0.28, y, 0.28, row_h, T.accent_rgb)
        if acc:
            gradient_fill(acc, T.accent_grad1, T.accent_grad2, 90)
            set_solid_alpha(acc, max(18, 56-i*7))
        nd = min(0.6, row_h*0.7)
        number_badge(slide, CX+CW-1.6, y+(row_h-nd)/2, nd, i+1, T)
        txt(slide, result, CX+0.25, y+0.07, CW-2.3, row_h-0.14,
            font=_FONT, size=max(10, min(12.5, row_h*8)),
            bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)

    slide_number(slide, 8, getattr(req, "_total_slides", 13), T)
    return slide

# ─── CONCLUSION ────────────────────────────────────────────────────────
def make_conclusion(prs, req: PresentationRequest, T: Theme):
    slide = _section_slide(prs, T, "💡", "خاتمة", "البحث")
    CY = 0.28; CH = H - CY - 0.22; cw = CW

    cc = rrect(slide, CX, CY, cw, CH, T.card_rgb, radius_pct=14)
    if cc:
        multi_stop_gradient(cc,[(0,T.card),(50,T.bg2),(100,T.card)],135)
        shadow(cc, blur=22, dist=6, alpha=0.44)
        glow(cc, T.accent.lstrip('#'), radius=26, alpha=0.08)

    tp = rrect(slide, CX, CY, cw, 0.3, T.accent_rgb, radius_pct=0)
    if tp:
        multi_stop_gradient(tp,[(0,T.accent2),(50,T.accent),(100,T.accent2)],0)
        glow(tp, T.accent.lstrip('#'), radius=11, alpha=0.28)

    diamond(slide, CX+0.28, CY+0.44, 0.95, 0.95, T.accent_rgb, alpha=13)

    txt(slide, "❝", CX+0.32, CY+0.42, 1.55, 1.4,
        font="Calibri", size=44, bold=False,
        color=T.accent_rgb, align=PP_ALIGN.LEFT, rtl=False)

    txt(slide, req.general_conclusion,
        CX+0.32, CY+1.28, cw-0.9, CH-2.2,
        font=_FONT, size=max(11, min(14, int((CH-2.2)*5))), bold=False,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT, rtl=True, vcenter=True, line_spacing=1.35)

    ny = CY+CH-0.88
    hl = rect(slide, CX+cw*0.18, ny, cw*0.64, 0.06, T.accent_rgb)
    if hl: multi_stop_gradient(hl,[(0,T.bg2),(50,T.accent),(100,T.bg2)],0)
    txt(slide, req.student_name, CX, ny+0.12, cw, 0.68,
        font=_FONT, size=13, bold=True,
        color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=True)

    slide_number(slide, 9, getattr(req, "_total_slides", 13), T)
    return slide


# ─── RECOMMENDATIONS ───────────────────────────────────────────────────
def make_recommendations(prs, req: PresentationRequest, T: Theme):
    slide = _section_slide(prs, T, "✅", "توصيات", "البحث")
    CY = 0.35; CH = H - CY - 0.3
    recs = req.recommendations[:8]
    n = len(recs)
    if not recs: return slide
    gap = 0.14
    row_h = (CH - gap*(n-1)) / n
    for i, rec in enumerate(recs):
        y = CY+i*(row_h+gap)
        rw = rrect(slide, CX, y, CW, row_h, T.card_rgb, radius_pct=9)
        if rw:
            multi_stop_gradient(rw,[(0,T.card),(100,T.bg2)],0)
            shadow(rw, blur=5, dist=2, alpha=0.2)
        dot = oval(slide, CX+CW-1.75, y+(row_h-0.35)/2, 0.35, 0.35, T.accent_rgb)
        if dot:
            multi_stop_gradient(dot,[(0,T.accent),(100,T.accent2)],135)
        acc = rect(slide, CX+CW-0.24, y, 0.24, row_h, T.accent_rgb)
        if acc: gradient_fill(acc, T.accent_grad1, T.accent_grad2, 90)
        txt(slide, rec, CX+0.25, y+0.07, CW-2.3, row_h-0.14,
            font=_FONT, size=max(10, min(12.5, row_h*8)),
            bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)
    slide_number(slide, 10, getattr(req, "_total_slides", 13), T)
    return slide

# ─── FUTURE ────────────────────────────────────────────────────────────
def make_future(prs, req: PresentationRequest, T: Theme):
    slide = _section_slide(prs, T, "🔭", "آفاق", "مستقبلية")
    CY = 0.35; CH = H - CY - 0.3
    items = req.future_work[:6]
    if not items: return slide
    cols = 2 if len(items)>3 else 1
    rows_n = (len(items)+cols-1)//cols
    gh = 0.22; gv = 0.18
    cw = (CW - gh*(cols-1)) / cols
    ch = (CH - gv*(rows_n-1)) / rows_n
    for i, item in enumerate(items):
        ci = i%cols; ri = i//cols
        x = CX+ci*(cw+gh); y = CY+ri*(ch+gv)
        cc = rrect(slide, x, y, cw, ch, T.card_rgb, radius_pct=11)
        if cc:
            multi_stop_gradient(cc,[(0,T.card),(70,T.bg2),(100,T.bg)],155)
            shadow(cc, blur=12, dist=3, alpha=0.32)
        tp = rrect(slide, x, y, cw, 0.26, T.accent_rgb, radius_pct=0)
        if tp: multi_stop_gradient(tp,[(0,T.accent),(100,T.accent2)],0)
        nd = min(0.82, ch*0.3)
        number_badge(slide, x+cw/2-nd/2, y+0.38, nd, i+1, T)
        hl = rect(slide, x+cw*0.18, y+nd+0.5, cw*0.64, 0.04, T.muted_rgb)
        txt(slide, item, x+0.25, y+nd+0.65, cw-0.5, ch-nd-0.82,
            font=_FONT, size=max(9.5, min(12, ch*5.5)),
            bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.CENTER, rtl=True)
    slide_number(slide, 11, getattr(req, "_total_slides", 13), T)
    return slide

# ─── REFERENCES ────────────────────────────────────────────────────────
def make_references(prs, req: PresentationRequest, T: Theme):
    slide = _section_slide(prs, T, "📚", "المراجع", "والمصادر")
    CY = 0.35; CH = H - CY - 0.3
    refs = req.references[:12]
    n = len(refs)
    if not refs: return slide
    gap = 0.1
    row_h = min((CH - gap*(n-1)) / n, 1.35)
    total_h = n*(row_h+gap)-gap
    sy = CY + (CH-total_h)/2
    for i, ref in enumerate(refs):
        y = sy+i*(row_h+gap)
        if y+row_h > H-0.18: break
        even = i%2==0
        rw = rrect(slide, CX, y, CW, row_h,
                   T.card_rgb if even else T.bg2_rgb, radius_pct=5)
        if rw:
            stops = [(0,T.card),(100,T.bg2)] if even else [(0,T.bg2),(100,T.card)]
            multi_stop_gradient(rw, stops, 0)
        acc = rect(slide, CX+CW-0.22, y, 0.22, row_h, T.accent_rgb)
        if acc: set_solid_alpha(acc, 52)
        nb = rrect(slide, CX+0.1, y+(row_h-0.36)/2, 0.68, 0.36, T.bg_rgb, radius_pct=40)
        if nb: set_solid_alpha(nb, 62)
        txt(slide, f"[{i+1}]", CX+0.1, y+(row_h-0.36)/2, 0.68, 0.36,
            font="Calibri", size=8.5, bold=True,
            color=T.accent_rgb, align=PP_ALIGN.CENTER, rtl=False)
        txt(slide, ref, CX+0.9, y+0.04, CW-1.4, row_h-0.08,
            font=_FONT, size=max(9, min(11, row_h*7.5)),
            bold=False, color=T.text_light_rgb,
            align=PP_ALIGN.RIGHT, rtl=True)
    slide_number(slide, 12, getattr(req, "_total_slides", 13), T)
    return slide

# ─── FINAL ─────────────────────────────────────────────────────────────
def make_final(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs); bg(slide, T.bg_rgb)
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,angle=135)

    # الشريط الجانبي
    gradient_rect(slide,0,0,SW+0.5,H,T.grad2,T.grad1,angle=180)
    sepc=rect(slide,SW+0.5,0,0.12,H,T.accent_rgb)
    if sepc: gradient_fill(sepc,T.accent_grad1,T.accent_grad2,90)
    oval(slide,-3,-3,9,9,T.accent_rgb,alpha=6)
    oval(slide,0.5,H-5.5,7,7,T.bg2_rgb,alpha=40)
    decorative_dots(slide,0.4,H*0.55,4,3,0.14,0.32,T.accent_rgb,alpha=11)

    txt(slide,"✦",SW/2-1.0,H*0.28,2.0,2.0,
        font="Calibri",size=28,bold=False,color=T.accent_rgb,
        align=PP_ALIGN.CENTER,rtl=False,vcenter=True)
    if req.institution:
        txt(slide,req.institution,0.2,H*0.52,SW+0.28,1.2,
            font=_FONT,size=10,bold=False,color=T.muted_rgb,
            align=PP_ALIGN.CENTER,rtl=True,vcenter=True)
    if req.year:
        yb=rrect(slide,0.4,H-1.42,SW-0.32,0.64,T.accent_rgb,radius_pct=40)
        if yb: multi_stop_gradient(yb,[(0,T.accent),(100,T.accent2)],0)
        txt(slide,req.year,0.4,H-1.42,SW-0.32,0.64,
            font="Calibri",size=13,bold=True,color=T.text_dark_rgb,
            align=PP_ALIGN.CENTER,rtl=False,vcenter=True)

    # منطقة المحتوى
    gradient_rect(slide,SW+0.62,0,W-SW-0.62,H,T.grad1,T.bg2,angle=135)
    oval(slide,W-8,-3,11,11,T.accent_rgb,alpha=4)
    diamond(slide,W-5,H-5,4,4,T.accent_rgb,alpha=5)

    mcx=SW+1.0; mcw=W-mcx-0.8
    # البطاقة تملأ 90% من الارتفاع
    ccy=H*0.05; cch=H*0.9

    cc=rrect(slide,mcx,ccy,mcw,cch,T.card_rgb,radius_pct=16)
    if cc:
        multi_stop_gradient(cc,[(0,T.card),(50,T.bg2),(100,T.card)],135)
        shadow(cc,blur=28,dist=8,alpha=0.5)
        glow(cc,T.accent.lstrip('#'),radius=34,alpha=0.1)

    tp=rrect(slide,mcx,ccy,mcw,0.4,T.accent_rgb,radius_pct=0)
    if tp:
        multi_stop_gradient(tp,[(0,T.bg),(30,T.accent2),(50,T.accent),(70,T.accent2),(100,T.bg)],0)
        glow(tp,T.accent.lstrip('#'),radius=16,alpha=0.34)

    bp=rrect(slide,mcx,ccy+cch-0.26,mcw,0.26,T.accent_rgb,radius_pct=0)
    if bp: set_solid_alpha(bp,45)

    # ✦ نجمة
    txt(slide,"✦",mcx+mcw/2-0.72,ccy+0.5,1.44,1.2,
        font="Calibri",size=22,bold=False,color=T.accent_rgb,
        align=PP_ALIGN.CENTER,rtl=False,vcenter=True)

    # شكراً وتقديراً
    txt(slide,"شكراً وتقديراً",mcx+0.7,ccy+1.05,mcw-1.4,cch*0.28,
        font=_FONT,size=34,bold=True,color=T.text_light_rgb,
        align=PP_ALIGN.CENTER,rtl=True,vcenter=True,line_spacing=1.1)

    # فاصل
    d1=rect(slide,mcx+mcw*0.14,ccy+cch*0.38,mcw*0.72,0.06,T.accent_rgb)
    if d1: multi_stop_gradient(d1,[(0,T.bg2),(50,T.accent),(100,T.bg2)],0)
    rect(slide,mcx+mcw*0.24,ccy+cch*0.38+0.14,mcw*0.52,0.03,T.muted_rgb)

    # اسم الطالب
    txt(slide,req.student_name,mcx+0.7,ccy+cch*0.42,mcw-1.4,cch*0.15,
        font=_FONT,size=20,bold=True,color=T.accent_rgb,
        align=PP_ALIGN.CENTER,rtl=True,vcenter=True)

    # عنوان المذكرة
    ts=req.title_ar[:65]+("..." if len(req.title_ar)>65 else "")
    txt(slide,ts,mcx+1.0,ccy+cch*0.58,mcw-2.0,cch*0.2,
        font=_FONT,size=11.5,bold=False,italic=True,color=T.muted_rgb,
        align=PP_ALIGN.CENTER,rtl=True,vcenter=True,line_spacing=1.25)

    # فاصل أسفل
    footer=[]
    if req.institution: footer.append(req.institution)
    if req.year: footer.append(req.year)
    if footer:
        fb=rrect(slide,mcx+mcw*0.1,ccy+cch*0.82,mcw*0.8,0.62,T.bg_rgb,radius_pct=40)
        if fb: set_solid_alpha(fb,50)
        txt(slide,"  ·  ".join(footer),mcx+0.8,ccy+cch*0.82,mcw-1.6,0.62,
            font=_FONT,size=11,bold=False,color=T.muted_rgb,
            align=PP_ALIGN.CENTER,rtl=True,vcenter=True)

    bt=rect(slide,0,H-0.28,W,0.28,T.accent_rgb)
    if bt: multi_stop_gradient(bt,[(0,T.bg),(30,T.accent),(70,T.accent2),(100,T.bg)],0)
    return slide

def make_intro(prs, req, T):
    slide = _section_slide(prs, T, "📖", "مقدمة", "البحث")
    CY=0.28; CH=H-CY-0.22
    items=[]
    if req.intro_overview: items.append(("📖","نظرة عامة",req.intro_overview))
    if req.intro_approach:  items.append(("🔬","المنهج المتبع",req.intro_approach))
    if not items: return slide
    n=len(items); gap=0.25
    cw=(CW-gap*(n-1))/n
    # ارتفاع البطاقة محدود لضمان بقاء النص بداخلها
    CARD_H = min(CH * 0.76, 10.5)
    card_y = CY + (CH - CARD_H) / 2
    ic_s = min(1.7, CARD_H * 0.22)
    ic_y_off = 0.48
    lbl_y_off = ic_y_off + ic_s + 0.26
    div_y_off = lbl_y_off + 0.68 + 0.1
    txt_y_off = div_y_off + 0.08
    txt_h = CARD_H - txt_y_off - 0.38

    for i,(icon,lbl,val) in enumerate(items[:2]):
        x=CX+i*(cw+gap)
        cc=rrect(slide,x,card_y,cw,CARD_H,T.card_rgb,radius_pct=12)
        if cc:
            multi_stop_gradient(cc,[(0,T.card),(100,T.card)],150)
            shadow(cc,blur=16,dist=5,alpha=0.48)
        tp=rrect(slide,x,card_y,cw,0.28,T.accent_rgb,radius_pct=0)
        if tp: multi_stop_gradient(tp,[(0,T.accent),(100,T.accent2)],0)
        icon_circle(slide,x+cw/2-ic_s/2,card_y+ic_y_off,ic_s,
                    T.accent_grad1,T.accent_grad2,icon,max(14,int(ic_s*11)),T)
        txt(slide,lbl,x+0.22,card_y+lbl_y_off,cw-0.44,0.68,
            font=_FONT,size=14,bold=True,color=T.accent_rgb,
            align=PP_ALIGN.CENTER,rtl=True,vcenter=True)
        rect(slide,x+cw*0.14,card_y+div_y_off,cw*0.72,0.04,T.muted_rgb)
        txt(slide,val,x+0.22,card_y+txt_y_off,cw-0.44,txt_h,
            font=_FONT,size=max(10,min(12,txt_h*2.2)),bold=False,
            color=T.text_light_rgb,align=PP_ALIGN.RIGHT,
            rtl=True,vcenter=True,line_spacing=1.3)
    slide_number(slide,1,13,T)
    return slide
