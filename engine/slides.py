"""
Canva Engine v17.4 — نصوص احترافية مع توسيط عمودي كامل
"""
from __future__ import annotations
from pptx.enum.text import PP_ALIGN
from engine.primitives import (
    W, H, rect, rrect, oval, bg, hline, vline,
    gradient_fill, gradient_rect, shadow, set_solid_alpha,
    multi_stop_gradient, glow, diamond, hexagon, decorative_dots,
    card_3d, icon_circle, number_badge, slide_number,
    txt, txt2, blank_slide,
)
from core.themes import Theme
from core.models import PresentationRequest

_FONT = "Cairo"
def set_font(n): global _FONT; _FONT = n

HEADER_H = 2.9

# ── خلفيات ─────────────────────────────────────────────────────────────
def _bg(slide, T, style='a'):
    bg(slide, T.bg_rgb)
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,
                  angle={'a':135,'b':160,'c':90,'d':45}[style])
    if style=='a':
        oval(slide,-3,-3,11,11,T.accent_rgb,alpha=5)
        oval(slide,W-9,H-8,14,14,T.bg2_rgb,alpha=45)
        decorative_dots(slide,1.2,H-4.2,5,3,0.16,0.42,T.accent_rgb,alpha=12)
    elif style=='b':
        diamond(slide,W-7,-2,6,6,T.accent_rgb,alpha=6)
        diamond(slide,-1.5,H-4.5,4.5,4.5,T.accent_rgb,alpha=5)
        hexagon(slide,W-4.5,H*0.3,2.8,2.8,T.accent_rgb,alpha=7)
        decorative_dots(slide,1.0,1.8,4,4,0.15,0.36,T.accent_rgb,alpha=9)
    elif style=='c':
        oval(slide,-4,-3,12,12,T.accent_rgb,alpha=4)
        oval(slide,W-10,H-9,15,15,T.accent_rgb,alpha=4)
        oval(slide,W-6,-2,9,9,T.bg2_rgb,alpha=40)
        decorative_dots(slide,W-6.5,1.5,4,5,0.14,0.35,T.accent_rgb,alpha=10)
    elif style=='d':
        for r,a in [(26,4),(20,5),(14,6),(8,8)]:
            oval(slide,W/2-r/2,H/2-r/2,r,r,T.accent_rgb,alpha=a)
        decorative_dots(slide,1.8,H-3.8,5,2,0.18,0.44,T.accent_rgb,alpha=12)

# ── هيدر ──────────────────────────────────────────────────────────────
def _hdr(slide, T, title, sub='', side='right', slide_num=None, total_slides=13, req=None):
    if req is not None and hasattr(req, '_total_slides'):
        total_slides = req._total_slides
    gradient_rect(slide,0,0,W,HEADER_H,T.grad2,T.grad1,angle=135)
    al = rect(slide,0,HEADER_H-0.22,W,0.22,T.accent_rgb)
    if al: multi_stop_gradient(al,[(0,T.bg),(40,T.accent),(60,T.accent2),(100,T.bg)],0)
    if side=='right':
        av = rect(slide,W-0.52,0,0.52,HEADER_H,T.accent_rgb)
        if av: gradient_fill(av,T.accent_grad1,T.accent_grad2,90)
    else:
        av = rect(slide,0,0,0.52,HEADER_H,T.accent_rgb)
        if av: gradient_fill(av,T.accent_grad1,T.accent_grad2,90)
    oval(slide,W-5,-2,7,7,T.accent_rgb,alpha=9)

    # رقم الشريحة داخل الهيدر — أعلى اليسار، منفصل تماماً عن العنوان
    if slide_num is not None:
        nb_s = 0.72
        nb_x = 1.1
        nb_y = (HEADER_H - nb_s) / 2
        nb = oval(slide, nb_x, nb_y, nb_s, nb_s, T.accent_rgb)
        if nb:
            from engine.primitives import gradient_fill as _gf, shadow as _sh
            _gf(nb, T.accent_grad1, T.accent_grad2, 135)
            _sh(nb, blur=8, dist=2, alpha=0.35)
        txt(slide, str(slide_num), nb_x, nb_y, nb_s, nb_s,
            font="Calibri", size=14, bold=True,
            color=T.text_dark_rgb, align=PP_ALIGN.CENTER, rtl=False, vcenter=True)
        # نص المجموع صغير
        txt(slide, f"/{total_slides}", nb_x + nb_s, nb_y + nb_s*0.3, 0.8, nb_s*0.4,
            font="Calibri", size=8, bold=False,
            color=T.muted_rgb, align=PP_ALIGN.LEFT, rtl=False, vcenter=True)
        title_x = nb_x + nb_s + 1.0
    else:
        title_x = 0.7

    # عنوان الشريحة — يبدأ بعد رقم الشريحة دون تداخل
    title_w = W - title_x - 0.8
    txt(slide, title, title_x, 0.2, title_w, HEADER_H*0.62,
        font=_FONT, size=26, bold=True,
        color=T.text_light_rgb, align=PP_ALIGN.RIGHT,
        rtl=True, vcenter=True, line_spacing=1.1)

    # عنوان فرعي — أصغر وأشد خفوتاً
    if sub:
        txt(slide, sub, title_x, HEADER_H*0.62, title_w, HEADER_H*0.35,
            font=_FONT, size=13, bold=False, italic=True,
            color=T.muted_rgb, align=PP_ALIGN.RIGHT,
            rtl=True, vcenter=True, line_spacing=1.0)

# ── محتوى Y وH ─────────────────────────────────────────────────────────
CY0 = HEADER_H + 0.3
def _ch(): return H - CY0 - 0.28

# ══════════════════════════════════════════════════════════════════════
# COVER
# ══════════════════════════════════════════════════════════════════════
def make_cover(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs)
    _bg(slide,T,'b')

    tp=rect(slide,0,0,W,0.44,T.accent_rgb)
    if tp: multi_stop_gradient(tp,[(0,T.bg),(30,T.accent),(70,T.accent2),(100,T.bg)],0)
    bt=rect(slide,0,H-0.32,W,0.32,T.accent_rgb)
    if bt: gradient_fill(bt,T.accent_grad1,T.accent_grad2,0)

    if req.institution:
        ib=rrect(slide,W/2-10,0.5,20,0.7,T.card_rgb,radius_pct=40)
        if ib: set_solid_alpha(ib,75)
        txt(slide,req.institution,W/2-10,0.5,20,0.7,
            font=_FONT,size=11,bold=False,color=T.muted_rgb,
            align=PP_ALIGN.CENTER,rtl=True,vcenter=True)

    # عنوان بارتفاع ثابت مناسب للنص + معلومات تملأ الباقي
    title_y=1.35; title_h=7.5
    info_y=title_y+title_h+0.18
    info_h=H-0.36-info_y-0.08
    cx=1.8; cw=W-3.6

    mc=rrect(slide,cx,title_y,cw,title_h,T.card_rgb,radius_pct=14)
    if mc:
        multi_stop_gradient(mc,[(0,T.card),(100,T.bg2)],angle=135)
        shadow(mc,blur=24,dist=7,alpha=0.48)
    ct=rrect(slide,cx,title_y,cw,0.36,T.accent_rgb,radius_pct=0)
    if ct: multi_stop_gradient(ct,[(0,T.accent),(50,T.accent2),(100,T.accent)],0)
    vline(slide,cx+cw-0.22,title_y+0.36,title_h-0.36,T.accent_rgb,thickness=0.22)

    ts=24 if len(req.title_ar)<42 else 19 if len(req.title_ar)<65 else 16
    txt(slide,req.title_ar,cx+0.45,title_y+0.38,cw-0.9,title_h*0.7,
        font=_FONT,size=ts,bold=True,color=T.text_light_rgb,
        align=PP_ALIGN.CENTER,rtl=True,vcenter=True,line_spacing=1.2)
    if req.title_en:
        txt(slide,req.title_en,cx+0.45,title_y+title_h*0.67,cw-0.9,title_h*0.2,
            font="Calibri",size=10.5,bold=False,italic=True,
            color=T.muted_rgb,align=PP_ALIGN.CENTER,rtl=False,vcenter=True)
    hl=rect(slide,cx+cw*0.1,title_y+title_h*0.88,cw*0.8,0.05,T.accent_rgb)
    if hl: multi_stop_gradient(hl,[(0,T.bg2),(50,T.accent),(100,T.bg2)],0)

    ic=rrect(slide,cx,info_y,cw,info_h,T.card_rgb,radius_pct=12)
    if ic:
        multi_stop_gradient(ic,[(0,T.bg2),(100,T.card)],135)
        shadow(ic,blur=14,dist=4,alpha=0.34)
    vline(slide,cx+cw-0.18,info_y,info_h,T.accent_rgb,thickness=0.18)

    rows=[("الطالب",req.student_name)]
    if req.supervisor:     rows.append(("المشرف",req.supervisor))
    if req.co_supervisor:  rows.append(("المشرف المساعد",req.co_supervisor))
    if req.specialization: rows.append(("التخصص",req.specialization))
    if req.year:           rows.append(("السنة",req.year))

    rh=info_h/max(len(rows),1)
    for i,(lbl,val) in enumerate(rows):
        y=info_y+i*rh
        rb=rrect(slide,cx+0.25,y+0.04,cw-0.62,rh-0.08,T.bg_rgb,radius_pct=7)
        if rb: set_solid_alpha(rb,50)
        txt(slide,f"{lbl} :",cx+0.42,y+0.04,4.5,rh-0.08,
            font=_FONT,size=max(10.5,min(12.5,rh*7.5)),bold=True,
            color=T.accent_rgb,align=PP_ALIGN.RIGHT,rtl=True,vcenter=True)
        vline(slide,cx+5.15,y+rh*0.12,rh*0.76,T.muted_rgb,thickness=0.04)
        txt(slide,val,cx+5.35,y+0.04,cw-6.0,rh-0.08,
            font=_FONT,size=max(12,min(14.5,rh*9)),bold=False,
            color=T.text_light_rgb,align=PP_ALIGN.RIGHT,rtl=True,vcenter=True)
    return slide

# ══════════════════════════════════════════════════════════════════════
# INTRO
# ══════════════════════════════════════════════════════════════════════
def make_intro(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs); _bg(slide,T,'c')
    _hdr(slide,T,"مقدمة البحث","نظرة عامة على الدراسة",'right',slide_num=1,req=req)
    CY=CY0; CH=_ch()
    items=[]
    if req.intro_overview: items.append(("📖","نظرة عامة",req.intro_overview))
    if req.intro_approach:  items.append(("🔬","المنهج المتبع",req.intro_approach))
    if not items: return slide
    n=len(items); gap=0.4
    col_w=(W-2.4-gap*(n-1))/n
    # ارتفاع البطاقة محدود بحيث يبقى المحتوى ضمنها
    CARD_H = min(CH * 0.72, 10.5)
    card_y = CY + (CH - CARD_H) / 2  # توسيط عمودي للبطاقة
    ic_s = 1.8
    lbl_h = 0.75
    div_h = 0.08
    padding_top = 0.5   # مسافة من أعلى البطاقة للدائرة
    ic_y_offset = padding_top
    lbl_y_offset = ic_y_offset + ic_s + 0.28
    div_y_offset = lbl_y_offset + lbl_h + 0.1
    txt_y_offset = div_y_offset + div_h + 0.15
    txt_h = CARD_H - txt_y_offset - 0.4  # يبقى النص ضمن البطاقة مع هامش 0.4 أسفلها

    for i,(icon,lbl,val) in enumerate(items[:2]):
        x=1.2+i*(col_w+gap)
        cc=rrect(slide,x,card_y,col_w,CARD_H,T.card_rgb,radius_pct=10)
        if cc:
            multi_stop_gradient(cc,[(0,T.card),(100,T.card)],150)
            shadow(cc,blur=20,dist=7,alpha=0.5)
        # شريط علوي ملوّن
        tp=rrect(slide,x,card_y,col_w,0.32,T.accent_rgb,radius_pct=0)
        if tp: multi_stop_gradient(tp,[(0,T.accent),(100,T.accent2)],0)
        # دائرة الأيقونة — داخل البطاقة
        ic_x = x + col_w/2 - ic_s/2
        ic_y = card_y + ic_y_offset
        icon_circle(slide,ic_x,ic_y,ic_s,
                    T.accent_grad1,T.accent_grad2,icon,max(16,int(ic_s*11)),T)
        # عنوان القسم
        txt(slide,lbl,x+0.22,card_y+lbl_y_offset,col_w-0.44,lbl_h,
            font=_FONT,size=16,bold=True,color=T.accent_rgb,
            align=PP_ALIGN.CENTER,rtl=True,vcenter=True)
        # خط فاصل
        hline(slide,x+col_w*0.14,card_y+div_y_offset,col_w*0.72,T.accent_rgb,thickness=0.04)
        # المحتوى — داخل البطاقة بالكامل
        txt(slide,val,x+0.28,card_y+txt_y_offset,col_w-0.56,txt_h,
            font=_FONT,size=max(11,min(13,txt_h*2.2)),bold=False,
            color=T.text_light_rgb,align=PP_ALIGN.RIGHT,
            rtl=True,vcenter=True,line_spacing=1.3)
    pass  # رقم الشريحة مدمج في الهيدر
    return slide

# ══════════════════════════════════════════════════════════════════════
# PLAN
# ══════════════════════════════════════════════════════════════════════
def make_plan(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs); _bg(slide,T,'a')
    _hdr(slide,T,"خطة البحث",f"يتضمن البحث {len(req.chapters)} فصول رئيسية",'left',slide_num=2,req=req)
    CY=CY0; CH=_ch()
    chapters=req.chapters[:8]; n=len(chapters)
    if not chapters: return slide
    gap=0.18
    row_h=(CH-gap*(n-1))/n

    for i,ch in enumerate(chapters):
        y=CY+i*(row_h+gap)
        even=i%2==0
        # الصف
        rw=rrect(slide,1.0,y,W-2.0,row_h,T.card_rgb if even else T.bg2_rgb,radius_pct=8)
        if rw:
            stops=[(0,T.card),(100,T.bg2)] if even else [(0,T.bg2),(100,T.card)]
            multi_stop_gradient(rw,stops,0)
            shadow(rw,blur=5,dist=2,alpha=0.18)
        # acc يميني
        acc=rect(slide,W-1.25,y,0.22,row_h,T.accent_rgb)
        if acc: gradient_fill(acc,T.accent_grad1,T.accent_grad2,90)
        # دائرة الرقم — متوسطة عمودياً
        nd=min(0.72,row_h*0.74)
        nx=W-2.52; ny=y+(row_h-nd)/2
        nc=oval(slide,nx,ny,nd,nd,T.accent_rgb)
        if nc:
            multi_stop_gradient(nc,[(0,T.accent),(100,T.accent2)],135)
            shadow(nc,blur=6,dist=2,alpha=0.28)
        txt(slide,str(i+1),nx,ny,nd,nd,
            font="Calibri",size=max(9,int(nd*11)),bold=True,
            color=T.text_dark_rgb,align=PP_ALIGN.CENTER,rtl=False,vcenter=True)
        # عنوان الفصل — محاذاة يمين مع توسيط عمودي
        txt(slide,ch.title,1.25,y,W-4.65,row_h,
            font=_FONT,size=max(11,min(15,int(row_h*9.5))),bold=False,
            color=T.text_light_rgb,align=PP_ALIGN.RIGHT,
            rtl=True,vcenter=True,line_spacing=1.15)
        # الصفحات
        if ch.pages:
            pb=rrect(slide,1.12,y+(row_h-0.36)/2,2.0,0.36,T.bg_rgb,radius_pct=40)
            if pb: set_solid_alpha(pb,55)
            txt(slide,ch.pages,1.12,y+(row_h-0.36)/2,2.0,0.36,
                font="Calibri",size=9,bold=False,color=T.muted_rgb,
                align=PP_ALIGN.CENTER,rtl=False,vcenter=True)
    pass  # رقم الشريحة مدمج في الهيدر
    return slide

# ══════════════════════════════════════════════════════════════════════
# PROBLEM
# ══════════════════════════════════════════════════════════════════════
def make_problem(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs); _bg(slide,T,'b')
    _hdr(slide,T,"إشكالية البحث","التساؤلات الرئيسية والفرعية",'right',slide_num=3,req=req)
    CY=CY0; CH=_ch()

    secs=[]; weights={}
    if req.main_problem:  secs.append('p'); weights['p']=2.6
    if req.main_question: secs.append('q'); weights['q']=1.6
    if req.sub_questions: secs.append('s'); weights['s']=2.0
    if not secs: return slide

    tw=sum(weights[s] for s in secs)
    gap=0.22; avail=CH-gap*(len(secs)-1)
    cy=CY

    if 'p' in secs:
        h=avail*weights['p']/tw
        pc=rrect(slide,1.0,cy,W-2.0,h,T.card_rgb,radius_pct=12)
        if pc:
            multi_stop_gradient(pc,[(0,T.card),(100,T.bg2)],135)
            shadow(pc,blur=18,dist=5,alpha=0.42)
            glow(pc,T.accent.lstrip('#'),radius=22,alpha=0.09)
        # تاج الإشكالية
        lb=rrect(slide,W-7.5,cy,5.8,0.52,T.accent_rgb,radius_pct=0)
        if lb: multi_stop_gradient(lb,[(0,T.accent),(100,T.accent2)],0)
        txt(slide,"◆  الإشكالية الرئيسية",W-7.5,cy,5.8,0.52,
            font=_FONT,size=11.5,bold=True,color=T.text_dark_rgb,
            align=PP_ALIGN.CENTER,rtl=True,vcenter=True)
        # علامة اقتباس كبيرة
        txt(slide,"❝",1.3,cy+0.6,1.4,1.1,
            font="Calibri",size=32,bold=False,color=T.accent_rgb,
            align=PP_ALIGN.LEFT,rtl=False,vcenter=False)
        # نص الإشكالية
        txt(slide,req.main_problem,2.85,cy+0.6,W-4.2,h-0.78,
            font=_FONT,size=max(11,min(13.5,h*5)),bold=False,
            color=T.text_light_rgb,align=PP_ALIGN.RIGHT,
            rtl=True,vcenter=True,line_spacing=1.35)
        cy+=h+gap

    if 'q' in secs:
        h=avail*weights['q']/tw
        qc=rrect(slide,1.0,cy,W-2.0,h,T.bg2_rgb,radius_pct=10)
        if qc: shadow(qc,blur=8,dist=2,alpha=0.25)
        vline(slide,W-1.4,cy,h,T.accent_rgb,thickness=0.22)
        dot=oval(slide,W-3.2,cy+h/2-0.2,0.4,0.4,T.accent_rgb)
        if dot: multi_stop_gradient(dot,[(0,T.accent),(100,T.accent2)],135)
        txt(slide,req.main_question,1.3,cy,W-3.6,h,
            font=_FONT,size=max(11,min(13.5,h*5.5)),bold=True,italic=True,
            color=T.text_light_rgb,align=PP_ALIGN.RIGHT,
            rtl=True,vcenter=True,line_spacing=1.2)
        cy+=h+gap

    if 's' in secs and req.sub_questions:
        h=avail*weights['s']/tw
        subs=req.sub_questions[:4]
        sub_h=h/len(subs)
        for i,q in enumerate(subs):
            y2=cy+i*sub_h
            if i%2==0:
                sb=rrect(slide,1.0,y2,W-2.0,sub_h-0.06,T.card_rgb,radius_pct=5)
                if sb: set_solid_alpha(sb,48)
            # رقم دائري صغير
            nd2=min(0.44,sub_h*0.55)
            nc2=oval(slide,W-2.8,y2+(sub_h-nd2)/2,nd2,nd2,T.accent_rgb)
            if nc2: set_solid_alpha(nc2,68)
            txt(slide,str(i+1),W-2.8,y2+(sub_h-nd2)/2,nd2,nd2,
                font="Calibri",size=max(7,int(nd2*9)),bold=True,
                color=T.accent_rgb,align=PP_ALIGN.CENTER,rtl=False,vcenter=True)
            txt(slide,q,1.3,y2,W-3.7,sub_h,
                font=_FONT,size=max(10,min(12.5,sub_h*7.5)),bold=False,
                color=T.muted_rgb,align=PP_ALIGN.RIGHT,
                rtl=True,vcenter=True,line_spacing=1.15)
    pass  # رقم الشريحة مدمج في الهيدر
    return slide

# ══════════════════════════════════════════════════════════════════════
# OBJECTIVES
# ══════════════════════════════════════════════════════════════════════
def make_objectives(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs); _bg(slide,T,'c')
    _hdr(slide,T,"أهداف البحث وفرضياته","",'left',slide_num=4,req=req)
    CY=CY0; CH=_ch()
    cols=[]
    if req.objectives: cols.append(("🎯  الأهداف",req.objectives))
    if req.hypotheses:  cols.append(("💡  الفرضيات",req.hypotheses))
    if not cols: return slide
    n_c=len(cols); gap=0.3
    col_w=(W-2.0-gap*(n_c-1))/n_c
    for i,(lbl,items) in enumerate(cols):
        x=1.0+i*(col_w+gap)
        cc=rrect(slide,x,CY,col_w,CH,T.card_rgb,radius_pct=12)
        if cc:
            multi_stop_gradient(cc,[(0,T.card),(100,T.bg2)],150)
            shadow(cc,blur=16,dist=5,alpha=0.38)
        # هيدر العمود
        hh=0.74
        hdr=rrect(slide,x,CY,col_w,hh,T.accent_rgb,radius_pct=0)
        if hdr: multi_stop_gradient(hdr,[(0,T.accent2),(100,T.accent)],0)
        txt(slide,lbl,x+0.18,CY,col_w-0.36,hh,
            font=_FONT,size=15,bold=True,color=T.text_dark_rgb,
            align=PP_ALIGN.CENTER,rtl=True,vcenter=True)
        # العناصر
        ia=CH-hh-0.12; n_items=min(len(items),8); ig=0.1
        ih=(ia-ig*(n_items-1))/n_items
        for j,item in enumerate(items[:8]):
            iy=CY+hh+0.06+j*(ih+ig)
            rb=rrect(slide,x+0.12,iy,col_w-0.24,ih,
                     T.bg2_rgb if j%2==0 else T.bg_rgb,radius_pct=7)
            if rb: set_solid_alpha(rb,72)
            # رقم
            number_badge(slide,x+col_w-0.82,iy+(ih-0.52)/2,0.52,j+1,T)
            # النص — توسيط عمودي كامل
            txt(slide,item,x+0.24,iy,col_w-1.26,ih,
                font=_FONT,size=max(9,min(12,ih*7.5)),bold=False,
                color=T.text_light_rgb,align=PP_ALIGN.RIGHT,
                rtl=True,vcenter=True,line_spacing=1.2)
    pass  # رقم الشريحة مدمج في الهيدر
    return slide

# ══════════════════════════════════════════════════════════════════════
# IMPORTANCE
# ══════════════════════════════════════════════════════════════════════
def make_importance(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs); _bg(slide,T,'b')
    _hdr(slide,T,"أهمية البحث","الأثر العلمي والعملي للدراسة",'right',slide_num=5,req=req)
    CY=CY0; CH=_ch()
    items=(req.importance or [])[:6]
    if not items: return slide
    icons=["⭐","🔑","📌","🌟","💎","🏆"]
    cols=2 if len(items)>3 else 1
    col_w=(W-2.0-0.28*(cols-1))/cols
    rows_n=(len(items)+cols-1)//cols
    gv=0.2
    card_h=(CH-gv*(rows_n-1))/rows_n
    for i,item in enumerate(items):
        ci=i%cols; ri=i//cols
        x=1.0+ci*(col_w+0.28); y=CY+ri*(card_h+gv)
        cc=card_3d(slide,x,y,col_w,card_h,T,radius=10)
        acc=rrect(slide,x+col_w-0.28,y,0.28,card_h,T.accent_rgb,radius_pct=0)
        if acc: multi_stop_gradient(acc,[(0,T.accent2),(100,T.accent)],90)
        ic_s=min(1.35,card_h*0.6)
        icon_circle(slide,x+0.28,y+(card_h-ic_s)/2,ic_s,
                    T.accent_grad1,T.accent_grad2,icons[i%len(icons)],
                    max(13,int(ic_s*11)),T)
        # النص مع توسيط عمودي
        txt(slide,item,x+ic_s+0.52,y+0.1,col_w-ic_s-1.05,card_h-0.2,
            font=_FONT,size=max(10,min(13,card_h*6.5)),bold=False,
            color=T.text_light_rgb,align=PP_ALIGN.RIGHT,
            rtl=True,vcenter=True,line_spacing=1.3)
    pass  # رقم الشريحة مدمج في الهيدر
    return slide

# ══════════════════════════════════════════════════════════════════════
# METHODOLOGY
# ══════════════════════════════════════════════════════════════════════
def make_methodology(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs); _bg(slide,T,'d')
    _hdr(slide,T,"منهجية البحث","الإجراءات والأدوات المستخدمة",'left',slide_num=6,req=req)
    CY=CY0; CH=_ch()
    icons_map={"المنهج":"📊","العينة":"👥","حجم العينة":"📏","الأداة":"🛠️"}
    fields=[]
    if req.methodology: fields.append(("المنهج",req.methodology))
    if req.sample_type:  fields.append(("العينة",req.sample_type))
    if req.sample_size:  fields.append(("حجم العينة",req.sample_size))
    if req.tool:         fields.append(("الأداة",req.tool))
    if not fields: return slide
    cols=2 if len(fields)>2 else len(fields)
    rows_n=(len(fields)+cols-1)//cols
    gh=0.28; gv=0.22
    col_w=(W-2.0-gh*(cols-1))/cols
    card_h=(CH-gv*(rows_n-1))/rows_n
    for i,(lbl,val) in enumerate(fields[:4]):
        ci=i%cols; ri=i//cols
        x=1.0+ci*(col_w+gh); y=CY+ri*(card_h+gv)
        cc=rrect(slide,x,y,col_w,card_h,T.card_rgb,radius_pct=12)
        if cc:
            multi_stop_gradient(cc,[(0,T.card),(100,T.bg2)],145)
            shadow(cc,blur=15,dist=4,alpha=0.4)
        ic_s=min(2.0,card_h*0.38)
        ic_x=x+col_w/2-ic_s/2
        ic_bg=oval(slide,ic_x,y+0.26,ic_s,ic_s,T.accent_rgb)
        if ic_bg:
            multi_stop_gradient(ic_bg,[(0,T.accent),(100,T.accent2)],135)
            shadow(ic_bg,blur=9,dist=3,alpha=0.3)
            glow(ic_bg,T.accent.lstrip('#'),radius=14,alpha=0.16)
        txt(slide,icons_map.get(lbl,"📌"),ic_x,y+0.3,ic_s,ic_s*0.86,
            font="Calibri",size=max(14,int(ic_s*10)),bold=False,
            color=T.text_dark_rgb,align=PP_ALIGN.CENTER,rtl=False,vcenter=True)
        # عنوان القسم
        lbl_y=y+ic_s+0.38
        txt(slide,lbl,x+0.22,lbl_y,col_w-0.44,0.68,
            font=_FONT,size=13.5,bold=True,color=T.accent_rgb,
            align=PP_ALIGN.CENTER,rtl=True,vcenter=True)
        hline(slide,x+col_w*0.15,lbl_y+0.7,col_w*0.7,T.muted_rgb,thickness=0.04)
        # القيمة
        txt(slide,val,x+0.22,lbl_y+0.8,col_w-0.44,card_h-lbl_y+y-0.92,
            font=_FONT,size=max(9.5,min(12,card_h*4.5)),bold=False,
            color=T.text_light_rgb,align=PP_ALIGN.CENTER,
            rtl=True,vcenter=True,line_spacing=1.25)
    pass  # رقم الشريحة مدمج في الهيدر
    return slide

# ══════════════════════════════════════════════════════════════════════
# STATS / KPI
# ══════════════════════════════════════════════════════════════════════
def make_stats(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs); _bg(slide,T,'a')
    _hdr(slide,T,"الأرقام والإحصاءات الرئيسية","",'right',slide_num=7,req=req)
    CY=CY0; CH=_ch()
    stats=req.stats[:6]
    if not stats: return slide
    cols=3 if len(stats)>=3 else len(stats)
    rows_n=(len(stats)+cols-1)//cols
    gh=0.28; gv=0.22
    col_w=(W-2.0-gh*(cols-1))/cols
    card_h=(CH-gv*(rows_n-1))/rows_n
    for i,stat in enumerate(stats):
        ci=i%cols; ri=i//cols
        x=1.0+ci*(col_w+gh); y=CY+ri*(card_h+gv)
        if y+card_h > H-0.2: break
        cc=rrect(slide,x,y,col_w,card_h,T.card_rgb,radius_pct=14)
        if cc:
            multi_stop_gradient(cc,[(0,T.bg2),(50,T.card),(100,T.bg2)],135)
            shadow(cc,blur=18,dist=5,alpha=0.42)
        tp=rrect(slide,x,y,col_w,0.3,T.accent_rgb,radius_pct=0)
        if tp:
            multi_stop_gradient(tp,[(0,T.accent2),(50,T.accent),(100,T.accent2)],0)
            glow(tp,T.accent.lstrip('#'),radius=9,alpha=0.25)
        bp=rrect(slide,x,y+card_h-0.2,col_w,0.2,T.accent_rgb,radius_pct=0)
        if bp: set_solid_alpha(bp,35)
        # القيمة — ضخمة في المنتصف
        vs=38 if len(stat.value)<=3 else 28 if len(stat.value)<=6 else 20
        txt(slide,stat.value,x+0.15,y+0.28,col_w-0.3,card_h*0.5,
            font="Calibri",size=vs,bold=True,color=T.accent_rgb,
            align=PP_ALIGN.CENTER,rtl=False,vcenter=True)
        # الوحدة
        if stat.unit:
            ub=rrect(slide,x+col_w/2-1.55,y+card_h*0.53+0.08,3.1,0.44,
                     T.bg_rgb,radius_pct=40)
            if ub: set_solid_alpha(ub,52)
            txt(slide,stat.unit,x+col_w/2-1.55,y+card_h*0.53+0.08,3.1,0.44,
                font=_FONT,size=10,bold=False,color=T.muted_rgb,
                align=PP_ALIGN.CENTER,rtl=True,vcenter=True)
        # فاصل
        hline(slide,x+col_w*0.14,y+card_h*0.7,col_w*0.72,T.muted_rgb,thickness=0.04)
        # التسمية — أسفل
        txt(slide,stat.label,x+0.15,y+card_h*0.72,col_w-0.3,card_h*0.26,
            font=_FONT,size=max(9,min(11,card_h*5)),bold=False,
            color=T.text_light_rgb,align=PP_ALIGN.CENTER,
            rtl=True,vcenter=True,line_spacing=1.1)
    pass  # رقم الشريحة مدمج في الهيدر
    return slide

# ══════════════════════════════════════════════════════════════════════
# RESULTS
# ══════════════════════════════════════════════════════════════════════
def make_results(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs); _bg(slide,T,'c')
    _hdr(slide,T,"نتائج البحث","أبرز ما توصلت إليه الدراسة",'left',slide_num=8,req=req)
    CY=CY0; CH=_ch()
    results=req.main_results[:8]; n=len(results)
    if not results: return slide
    gap=0.16; row_h=(CH-gap*(n-1))/n
    for i,result in enumerate(results):
        y=CY+i*(row_h+gap)
        even=i%2==0
        rw=rrect(slide,1.0,y,W-2.0,row_h,T.card_rgb if even else T.bg2_rgb,radius_pct=8)
        if rw:
            stops=[(0,T.card),(100,T.bg2)] if even else [(0,T.bg2),(100,T.card)]
            multi_stop_gradient(rw,stops,0)
            shadow(rw,blur=5,dist=2,alpha=0.18)
        acc=rect(slide,W-1.32,y,0.28,row_h,T.accent_rgb)
        if acc:
            gradient_fill(acc,T.accent_grad1,T.accent_grad2,90)
            set_solid_alpha(acc,max(18,56-i*7))
        nd=min(0.64,row_h*0.72)
        number_badge(slide,W-3.05,y+(row_h-nd)/2,nd,i+1,T)
        # النتيجة — توسيط عمودي كامل
        txt(slide,result,1.2,y,W-4.95,row_h,
            font=_FONT,size=max(10,min(13,row_h*7.5)),bold=False,
            color=T.text_light_rgb,align=PP_ALIGN.RIGHT,
            rtl=True,vcenter=True,line_spacing=1.25)
    pass  # رقم الشريحة مدمج في الهيدر
    return slide

# ══════════════════════════════════════════════════════════════════════
# CONCLUSION
# ══════════════════════════════════════════════════════════════════════
def make_conclusion(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs); _bg(slide,T,'d')
    _hdr(slide,T,"خاتمة البحث","الاستنتاج العام للدراسة",'right',slide_num=9,req=req)
    CY=CY0; CH=_ch(); cw=W-2.8
    cc=rrect(slide,1.4,CY,cw,CH,T.card_rgb,radius_pct=16)
    if cc:
        multi_stop_gradient(cc,[(0,T.card),(50,T.bg2),(100,T.card)],135)
        shadow(cc,blur=26,dist=7,alpha=0.48)
        glow(cc,T.accent.lstrip('#'),radius=28,alpha=0.09)
    tp=rrect(slide,1.4,CY,cw,0.36,T.accent_rgb,radius_pct=0)
    if tp:
        multi_stop_gradient(tp,[(0,T.accent2),(50,T.accent),(100,T.accent2)],0)
        glow(tp,T.accent.lstrip('#'),radius=13,alpha=0.32)
    diamond(slide,1.7,CY+0.5,1.0,1.0,T.accent_rgb,alpha=14)
    diamond(slide,W-2.8,CY+CH-1.7,0.9,0.9,T.accent_rgb,alpha=10)
    # علامة اقتباس
    txt(slide,"❝",2.0,CY+0.48,1.8,1.6,
        font="Calibri",size=48,bold=False,color=T.accent_rgb,
        align=PP_ALIGN.LEFT,rtl=False,vcenter=False)
    # الاستنتاج — يملأ البطاقة مع توسيط
    # الاستنتاج — يملأ البطاقة مع توسيط
    txt(slide,req.general_conclusion,2.0,CY+0.9,cw-1.2,CH-1.95,
        font=_FONT,size=max(12,min(15,CH*4.5)),bold=False,
        color=T.text_light_rgb,align=PP_ALIGN.RIGHT,
        rtl=True,vcenter=True,line_spacing=1.4)
    # الاسم
    ny=CY+CH-1.05
    hl=rect(slide,1.4+cw*0.18,ny,cw*0.64,0.06,T.accent_rgb)
    if hl: multi_stop_gradient(hl,[(0,T.bg2),(50,T.accent),(100,T.bg2)],0)
    txt(slide,req.student_name,1.4,ny+0.12,cw,0.75,
        font=_FONT,size=14,bold=True,color=T.accent_rgb,
        align=PP_ALIGN.CENTER,rtl=True,vcenter=True)
    pass  # رقم الشريحة مدمج في الهيدر
    return slide

# ══════════════════════════════════════════════════════════════════════
# RECOMMENDATIONS
# ══════════════════════════════════════════════════════════════════════
def make_recommendations(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs); _bg(slide,T,'b')
    _hdr(slide,T,"توصيات البحث","",'left',slide_num=10,req=req)
    CY=CY0; CH=_ch()
    recs=req.recommendations[:8]; n=len(recs)
    if not recs: return slide
    gap=0.16; row_h=(CH-gap*(n-1))/n
    for i,rec in enumerate(recs):
        y=CY+i*(row_h+gap)
        rw=rrect(slide,1.0,y,W-2.0,row_h,T.card_rgb,radius_pct=9)
        if rw:
            multi_stop_gradient(rw,[(0,T.card),(100,T.bg2)],0)
            shadow(rw,blur=6,dist=2,alpha=0.22)
        dot=oval(slide,W-1.92,y+(row_h-0.4)/2,0.4,0.4,T.accent_rgb)
        if dot:
            multi_stop_gradient(dot,[(0,T.accent),(100,T.accent2)],135)
            shadow(dot,blur=5,dist=1,alpha=0.28)
        acc=rect(slide,W-1.3,y,0.26,row_h,T.accent_rgb)
        if acc: gradient_fill(acc,T.accent_grad1,T.accent_grad2,90)
        txt(slide,rec,1.2,y,W-3.55,row_h,
            font=_FONT,size=max(10,min(13,row_h*7.5)),bold=False,
            color=T.text_light_rgb,align=PP_ALIGN.RIGHT,
            rtl=True,vcenter=True,line_spacing=1.25)
    pass  # رقم الشريحة مدمج في الهيدر
    return slide

# ══════════════════════════════════════════════════════════════════════
# FUTURE
# ══════════════════════════════════════════════════════════════════════
def make_future(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs); _bg(slide,T,'a')
    _hdr(slide,T,"آفاق البحث المستقبلية","",'right',slide_num=11,req=req)
    CY=CY0; CH=_ch()
    items=req.future_work[:6]
    if not items: return slide
    cols=2 if len(items)>3 else 1
    rows_n=(len(items)+cols-1)//cols
    gh=0.28; gv=0.22
    col_w=(W-2.0-gh*(cols-1))/cols
    card_h=(CH-gv*(rows_n-1))/rows_n
    for i,item in enumerate(items):
        ci=i%cols; ri=i//cols
        x=1.0+ci*(col_w+gh); y=CY+ri*(card_h+gv)
        cc=rrect(slide,x,y,col_w,card_h,T.card_rgb,radius_pct=12)
        if cc:
            multi_stop_gradient(cc,[(0,T.card),(70,T.bg2),(100,T.bg)],160)
            shadow(cc,blur=14,dist=4,alpha=0.36)
        tp=rrect(slide,x,y,col_w,0.28,T.accent_rgb,radius_pct=0)
        if tp: multi_stop_gradient(tp,[(0,T.accent),(100,T.accent2)],0)
        nd=min(1.0,card_h*0.32)
        number_badge(slide,x+col_w/2-nd/2,y+0.36,nd,i+1,T)
        hline(slide,x+col_w*0.18,y+nd+0.5,col_w*0.64,T.muted_rgb,thickness=0.04)
        txt(slide,item,x+0.3,y+nd+0.66,col_w-0.6,card_h-nd-0.84,
            font=_FONT,size=max(10,min(13,card_h*5)),bold=False,
            color=T.text_light_rgb,align=PP_ALIGN.CENTER,
            rtl=True,vcenter=True,line_spacing=1.3)
    pass  # رقم الشريحة مدمج في الهيدر
    return slide

# ══════════════════════════════════════════════════════════════════════
# REFERENCES
# ══════════════════════════════════════════════════════════════════════
def make_references(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs); _bg(slide,T,'c')
    _hdr(slide,T,"قائمة المراجع والمصادر","",'left',slide_num=12,req=req)
    CY=CY0; CH=_ch()
    refs=req.references[:12]; n=len(refs)
    if not refs: return slide
    gap=0.12; row_h=(CH-gap*(n-1))/n
    for i,ref in enumerate(refs):
        y=CY+i*(row_h+gap)
        if y+row_h > H-0.2: break
        even=i%2==0
        rw=rrect(slide,1.0,y,W-2.0,row_h,T.card_rgb if even else T.bg2_rgb,radius_pct=5)
        if rw:
            stops=[(0,T.card),(100,T.bg2)] if even else [(0,T.bg2),(100,T.card)]
            multi_stop_gradient(rw,stops,0)
        acc=rect(slide,W-1.28,y,0.24,row_h,T.accent_rgb)
        if acc: set_solid_alpha(acc,55)
        nb=rrect(slide,1.12,y+(row_h-0.38)/2,0.7,0.38,T.bg_rgb,radius_pct=40)
        if nb: set_solid_alpha(nb,65)
        txt(slide,f"[{i+1}]",1.12,y+(row_h-0.38)/2,0.7,0.38,
            font="Calibri",size=9,bold=True,color=T.accent_rgb,
            align=PP_ALIGN.CENTER,rtl=False,vcenter=True)
        txt(slide,ref,1.95,y+0.05,W-3.45,row_h-0.1,
            font=_FONT,size=max(9,min(11.5,row_h*7)),bold=False,
            color=T.text_light_rgb,align=PP_ALIGN.RIGHT,
            rtl=True,vcenter=True,line_spacing=1.15)
    pass  # رقم الشريحة مدمج في الهيدر
    return slide

# ══════════════════════════════════════════════════════════════════════
# FINAL
# ══════════════════════════════════════════════════════════════════════
def make_final(prs, req: PresentationRequest, T: Theme):
    slide = blank_slide(prs); bg(slide,T.bg_rgb)
    gradient_rect(slide,0,0,W,H,T.grad1,T.grad2,angle=135)
    oval(slide,-5,-5,15,15,T.accent_rgb,alpha=5)
    oval(slide,W-11,H-11,17,17,T.accent_rgb,alpha=4)
    oval(slide,W-6.5,-2.5,10,10,T.bg2_rgb,alpha=32)
    oval(slide,-2.5,H-6.5,9,9,T.bg2_rgb,alpha=28)
    diamond(slide,W*0.28,H*0.06,2.4,2.4,T.accent_rgb,alpha=8)
    diamond(slide,W*0.62,H*0.74,1.9,1.9,T.accent_rgb,alpha=6)
    decorative_dots(slide,1.4,H-4.8,7,3,0.17,0.4,T.accent_rgb,alpha=11)
    decorative_dots(slide,W-5.8,1.4,5,4,0.15,0.36,T.accent_rgb,alpha=9)

    cw=23; ch=11.5; cx=(W-cw)/2; cy=(H-ch)/2
    cc=rrect(slide,cx,cy,cw,ch,T.card_rgb,radius_pct=16)
    if cc:
        multi_stop_gradient(cc,[(0,T.card),(50,T.bg2),(100,T.card)],135)
        shadow(cc,blur=32,dist=10,alpha=0.55)
        glow(cc,T.accent.lstrip('#'),radius=38,alpha=0.12)
    tp=rrect(slide,cx,cy,cw,0.42,T.accent_rgb,radius_pct=0)
    if tp:
        multi_stop_gradient(tp,[(0,T.bg),(30,T.accent2),(50,T.accent),(70,T.accent2),(100,T.bg)],0)
        glow(tp,T.accent.lstrip('#'),radius=18,alpha=0.38)
    bp=rrect(slide,cx,cy+ch-0.28,cw,0.28,T.accent_rgb,radius_pct=0)
    if bp: set_solid_alpha(bp,48)

    txt(slide,"✦",cx+cw/2-0.75,cy+0.52,1.5,1.4,
        font="Calibri",size=26,bold=False,color=T.accent_rgb,
        align=PP_ALIGN.CENTER,rtl=False,vcenter=True)
    txt(slide,"شكراً وتقديراً",cx+0.8,cy+1.15,cw-1.6,2.7,
        font=_FONT,size=38,bold=True,color=T.text_light_rgb,
        align=PP_ALIGN.CENTER,rtl=True,vcenter=True,line_spacing=1.1)

    d1=rect(slide,cx+cw*0.14,cy+4.1,cw*0.72,0.06,T.accent_rgb)
    if d1: multi_stop_gradient(d1,[(0,T.bg2),(50,T.accent),(100,T.bg2)],0)
    rect(slide,cx+cw*0.24,cy+4.22,cw*0.52,0.03,T.muted_rgb)

    txt(slide,req.student_name,cx+0.8,cy+4.38,cw-1.6,1.35,
        font=_FONT,size=22,bold=True,color=T.accent_rgb,
        align=PP_ALIGN.CENTER,rtl=True,vcenter=True)

    ts=req.title_ar[:72]+("..." if len(req.title_ar)>72 else "")
    txt(slide,ts,cx+1.2,cy+5.85,cw-2.4,2.1,
        font=_FONT,size=12,bold=False,italic=True,color=T.muted_rgb,
        align=PP_ALIGN.CENTER,rtl=True,vcenter=True,line_spacing=1.3)

    footer=[]
    if req.institution: footer.append(req.institution)
    if req.year: footer.append(req.year)
    if footer:
        fb=rrect(slide,cx+cw*0.1,cy+ch-1.28,cw*0.8,0.6,T.bg_rgb,radius_pct=40)
        if fb: set_solid_alpha(fb,52)
        txt(slide,"  ·  ".join(footer),cx+0.8,cy+ch-1.28,cw-1.6,0.6,
            font=_FONT,size=11,bold=False,color=T.muted_rgb,
            align=PP_ALIGN.CENTER,rtl=True,vcenter=True)

    bbar=rect(slide,0,H-0.26,W,0.26,T.accent_rgb)
    if bbar: multi_stop_gradient(bbar,[(0,T.bg),(30,T.accent),(70,T.accent2),(100,T.bg)],0)
    return slide
