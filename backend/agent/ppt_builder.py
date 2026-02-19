"""
ppt_builder.py
All 7 animation types + Morph transitions + 6 unique slide layouts.
Layouts are randomly varied per slide within each presentation.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree
import io, random

# ── Helpers ───────────────────────────────────────────────────────────────────

def rgb(hex_str):
    h = hex_str.lstrip("#")
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

def set_bg(slide, color):
    fill = slide.background.fill
    fill.solid(); fill.fore_color.rgb = color

def rect(slide, l, t, w, h, color):
    s = slide.shapes.add_shape(1, l, t, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = color
    s.line.fill.background(); return s

def txt(slide, text, l, t, w, h, fn="Calibri", sz=18,
        col=RGBColor(255,255,255), bold=False, italic=False,
        align=PP_ALIGN.LEFT, wrap=True):
    b = slide.shapes.add_textbox(l, t, w, h)
    tf = b.text_frame; tf.word_wrap = wrap
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = text
    r.font.name = fn; r.font.size = Pt(sz)
    r.font.color.rgb = col; r.font.bold = bold; r.font.italic = italic
    return b

def bpara(tf, text, fn, sz, col, sp=10):
    p = tf.add_paragraph(); p.alignment = PP_ALIGN.LEFT
    p.space_before = Pt(sp); r = p.add_run()
    r.text = text; r.font.name = fn
    r.font.size = Pt(sz); r.font.color.rgb = col
    return p

def notes(slide, text):
    if text:
        try: slide.notes_slide.notes_text_frame.text = text
        except: pass

W = Inches(13.33); H = Inches(7.5)

# ── Animation ID counter ──────────────────────────────────────────────────────

_aid = [10]
def nid(): v=_aid[0]; _aid[0]+=1; return v
def reset_ids(): _aid[0]=10

P="http://schemas.openxmlformats.org/presentationml/2006/main"

# ── Animation XML builders ────────────────────────────────────────────────────

def anim_fade(spid, delay, dur=500):
    i1,i2,i3 = nid(),nid(),nid()
    return f"""<p:par xmlns:p="{P}"><p:cTn id="{i1}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="{i1}" nodeType="withEffect"><p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="{i2}" dur="1" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set><p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{i3}" dur="{dur}"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>"""

def anim_appear(spid, delay):
    i1,i2 = nid(),nid()
    return f"""<p:par xmlns:p="{P}"><p:cTn id="{i1}" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" grpId="{i1}" nodeType="withEffect"><p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="{i2}" dur="1" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set></p:childTnLst></p:cTn></p:par>"""

def anim_fly(spid, delay, dur=600):
    i1,i2,i3,i4 = nid(),nid(),nid(),nid()
    return f"""<p:par xmlns:p="{P}"><p:cTn id="{i1}" presetID="2" presetClass="entr" presetSubtype="8" fill="hold" grpId="{i1}" nodeType="withEffect"><p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="{i2}" dur="1" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set><p:animMotion origin="layout" path="M 0 0.25 L 0 0" pathEditMode="relative" rAng="0" ptsTypes="auto"><p:cBhvr><p:cTn id="{i3}" dur="{dur}" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr></p:animMotion><p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{i4}" dur="{dur}"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>"""

def anim_float(spid, delay, dur=700):
    i1,i2,i3,i4 = nid(),nid(),nid(),nid()
    return f"""<p:par xmlns:p="{P}"><p:cTn id="{i1}" presetID="2" presetClass="entr" presetSubtype="4" fill="hold" grpId="{i1}" nodeType="withEffect"><p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="{i2}" dur="1" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set><p:animMotion origin="layout" path="M 0 -0.15 L 0 0" pathEditMode="relative" rAng="0" ptsTypes="auto"><p:cBhvr><p:cTn id="{i3}" dur="{dur}" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr></p:animMotion><p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{i4}" dur="{dur}"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>"""

def anim_zoom(spid, delay, dur=500):
    i1,i2,i3,i4 = nid(),nid(),nid(),nid()
    return f"""<p:par xmlns:p="{P}"><p:cTn id="{i1}" presetID="18" presetClass="entr" presetSubtype="0" fill="hold" grpId="{i1}" nodeType="withEffect"><p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="{i2}" dur="1" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set><p:animScale><p:cBhvr><p:cTn id="{i3}" dur="{dur}" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr><p:from x="10000" y="10000"/><p:to x="100000" y="100000"/></p:animScale><p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="{i4}" dur="{dur}"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>"""

def anim_grow(spid, delay, dur=600):
    i1,i2 = nid(),nid()
    return f"""<p:par xmlns:p="{P}"><p:cTn id="{i1}" presetID="150" presetClass="emph" presetSubtype="0" fill="hold" grpId="{i1}" nodeType="withEffect"><p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst><p:animScale><p:cBhvr calcmode="lin" valueType="num"><p:cTn id="{i2}" dur="{dur//2}" autoRev="1"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr><p:from x="100000" y="100000"/><p:to x="115000" y="115000"/></p:animScale></p:childTnLst></p:cTn></p:par>"""

def anim_wipe(spid, delay, dur=600):
    i1,i2,i3 = nid(),nid(),nid()
    return f"""<p:par xmlns:p="{P}"><p:cTn id="{i1}" presetID="21" presetClass="entr" presetSubtype="8" fill="hold" grpId="{i1}" nodeType="withEffect"><p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="{i2}" dur="1" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set><p:animEffect transition="in" filter="wipe(right)"><p:cBhvr><p:cTn id="{i3}" dur="{dur}"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>"""

def anim_spin(spid, delay, dur=700):
    i1,i2 = nid(),nid()
    return f"""<p:par xmlns:p="{P}"><p:cTn id="{i1}" presetID="156" presetClass="emph" presetSubtype="0" fill="hold" grpId="{i1}" nodeType="withEffect"><p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst><p:animRot by="5400000"><p:cBhvr calcmode="lin"><p:cTn id="{i2}" dur="{dur}" autoRev="0"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animRot></p:childTnLst></p:cTn></p:par>"""

def anim_path(spid, delay, dur=800):
    i1,i2,i3 = nid(),nid(),nid()
    return f"""<p:par xmlns:p="{P}"><p:cTn id="{i1}" presetID="0" presetClass="path" presetSubtype="0" fill="hold" grpId="{i1}" nodeType="withEffect"><p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="{i2}" dur="1" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set><p:animMotion origin="layout" path="M -0.5 0 L 0 0" pathEditMode="relative" rAng="0" ptsTypes="auto"><p:cBhvr calcmode="lin"><p:cTn id="{i3}" dur="{dur}" fill="hold"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr></p:animMotion></p:childTnLst></p:cTn></p:par>"""

ANIM_FNS = {
    "fade": anim_fade, "appear": anim_appear, "fly": anim_fly,
    "float": anim_float, "zoom": anim_zoom, "grow": anim_grow,
    "wipe": anim_wipe, "spin": anim_spin, "path": anim_path,
}

def inject(slide, se):
    """se: list of (shape, anim_type, delay_ms)"""
    if not se: return
    blocks = ""
    for shape, atype, delay in se:
        fn = ANIM_FNS.get(atype, anim_fade)
        blocks += fn(shape.shape_id, delay)
    r=nid(); s=nid(); c=nid()
    xml=f"""<p:timing xmlns:p="{P}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:tnLst><p:par><p:cTn id="{r}" dur="indefinite" restart="whenNotActive" nodeType="tmRoot"><p:childTnLst><p:par><p:cTn id="{s}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="{c}" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>{blocks}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst/></p:timing>"""
    try:
        el = slide._element
        ex = el.find(qn("p:timing"))
        if ex is not None: el.remove(ex)
        el.append(etree.fromstring(xml.encode("utf-8")))
    except Exception as e:
        print(f"[Anim] {e}")

def morph(slide):
    try:
        xml="""<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" spd="med" advTm="0"><p14:morph option="byObject"/></p:transition>"""
        el=etree.fromstring(xml)
        ex=slide._element.find(qn("p:transition"))
        if ex is not None: slide._element.remove(ex)
        t=slide._element.find(qn("p:timing"))
        if t is not None: slide._element.insert(list(slide._element).index(t),el)
        else: slide._element.append(el)
    except Exception as e:
        print(f"[Morph] {e}")

# ── 6 slide layout builders ───────────────────────────────────────────────────

def layout_hero(prs, data, th, idx):
    """Full-bleed hero: big title, accent bar, circle deco, detail text."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, rgb(th["bg"])); se=[]
    bar=rect(slide,Inches(0),Inches(0),Inches(0.5),H,rgb(th["accent"]))
    se.append((bar,"wipe",0))
    circ=slide.shapes.add_shape(9,Inches(9.0),Inches(-2.0),Inches(5.8),Inches(5.8))
    circ.fill.solid(); circ.fill.fore_color.rgb=rgb(th["accent2"]); circ.line.fill.background()
    se.append((circ,"zoom",80))
    circ2=slide.shapes.add_shape(9,Inches(10.5),Inches(4.5),Inches(3.5),Inches(3.5))
    circ2.fill.solid(); circ2.fill.fore_color.rgb=rgb(th["accent"]); circ2.line.fill.background()
    se.append((circ2,"zoom",140))
    n=txt(slide,f"{idx:02d}",Inches(0.75),Inches(0.3),Inches(1.2),Inches(0.5),th["font_body"],11,rgb(th["accent"]))
    se.append((n,"appear",200))
    t=txt(slide,data["title"],Inches(0.75),Inches(1.3),Inches(10.5),Inches(2.7),th["font_head"],52,rgb(th["title_color"]),bold=True)
    se.append((t,"float",280))
    s=txt(slide,data.get("subtitle",""),Inches(0.75),Inches(3.85),Inches(10),Inches(0.8),th["font_body"],20,rgb(th["accent"]),italic=True)
    se.append((s,"fade",480))
    d=txt(slide,data.get("detail",""),Inches(0.75),Inches(4.75),Inches(10.5),Inches(2.0),th["font_body"],14,rgb(th["muted_color"]),wrap=True)
    se.append((d,"fade",660))
    inject(slide,se); morph(slide); notes(slide,data.get("notes","")); return slide

def layout_split(prs, data, th, idx):
    """Left bullets + right detail card."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, rgb(th["bg"])); se=[]
    bar=rect(slide,Inches(0),Inches(0),W,Inches(0.07),rgb(th["accent"]))
    se.append((bar,"wipe",0))
    n=txt(slide,f"{idx:02d}",Inches(0.4),Inches(0.22),Inches(1),Inches(0.4),th["font_body"],10,rgb(th["accent"]))
    se.append((n,"appear",100))
    t=txt(slide,data["title"],Inches(0.4),Inches(0.5),Inches(7.5),Inches(1.2),th["font_head"],36,rgb(th["title_color"]),bold=True)
    se.append((t,"fly",200))
    s=txt(slide,data.get("subtitle",""),Inches(0.4),Inches(1.65),Inches(7),Inches(0.5),th["font_body"],13,rgb(th["accent"]),italic=True)
    se.append((s,"fade",370))
    btxb=slide.shapes.add_textbox(Inches(0.4),Inches(2.3),Inches(6.5),Inches(4.8))
    btf=btxb.text_frame; btf.word_wrap=True
    for i,pt in enumerate(data.get("points",[])):
        if i==0:
            p0=btf.paragraphs[0]; p0.alignment=PP_ALIGN.LEFT; r=p0.add_run(); r.text="▸  "+pt
            r.font.name=th["font_body"]; r.font.size=Pt(13); r.font.color.rgb=rgb(th["body_color"])
        else: bpara(btf,"▸  "+pt,th["font_body"],13,rgb(th["body_color"]))
    se.append((btxb,"fly",480))
    card=rect(slide,Inches(7.3),Inches(0.4),Inches(5.65),Inches(6.7),rgb(th["card_bg"]))
    se.append((card,"fade",270))
    cb=rect(slide,Inches(7.3),Inches(0.4),Inches(0.1),Inches(6.7),rgb(th["accent"]))
    se.append((cb,"wipe",320))
    lbl=txt(slide,"DEEP DIVE",Inches(7.6),Inches(0.78),Inches(5),Inches(0.38),th["font_body"],9,rgb(th["accent"]),bold=True)
    se.append((lbl,"appear",390))
    d=txt(slide,data.get("detail",""),Inches(7.6),Inches(1.28),Inches(5.1),Inches(5.65),th["font_body"],14,rgb(th["body_color"]),wrap=True)
    se.append((d,"fade",540))
    inject(slide,se); morph(slide); notes(slide,data.get("notes","")); return slide

def layout_grid(prs, data, th, idx):
    """2×2 card grid for 4 key points."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, rgb(th["bg"])); se=[]
    strip=rect(slide,Inches(0),Inches(0),W,Inches(2.1),rgb(th["header_bg"]))
    se.append((strip,"fade",0))
    n=txt(slide,f"{idx:02d}",Inches(0.4),Inches(0.2),Inches(1),Inches(0.4),th["font_body"],10,rgb(th["accent"]))
    se.append((n,"appear",80))
    t=txt(slide,data["title"],Inches(0.4),Inches(0.42),Inches(12.5),Inches(1.1),th["font_head"],38,rgb(th["title_color"]),bold=True)
    se.append((t,"float",160))
    s=txt(slide,data.get("subtitle",""),Inches(0.4),Inches(1.52),Inches(10),Inches(0.47),th["font_body"],13,rgb(th["accent"]),italic=True)
    se.append((s,"fade",330))
    pts=data.get("points",[""])
    pos=[(Inches(0.3),Inches(2.25)),(Inches(6.85),Inches(2.25)),(Inches(0.3),Inches(4.95)),(Inches(6.85),Inches(4.95))]
    cw,ch=Inches(6.25),Inches(2.45)
    for i,(pt,(cx,cy)) in enumerate(zip(pts[:4],pos)):
        d=420+i*130
        card=rect(slide,cx,cy,cw,ch,rgb(th["card_bg"])); se.append((card,"zoom",d))
        badge=rect(slide,cx+Inches(0.15),cy+Inches(0.13),Inches(0.58),Inches(0.58),rgb(th["accent"]))
        se.append((badge,"grow",d+50))
        bn=txt(slide,f"0{i+1}",cx+Inches(0.15),cy+Inches(0.13),Inches(0.58),Inches(0.58),th["font_body"],12,rgb(th["bg"]),bold=True,align=PP_ALIGN.CENTER)
        se.append((bn,"appear",d+50))
        pt_=txt(slide,pt,cx+Inches(0.88),cy+Inches(0.1),cw-Inches(1.05),ch-Inches(0.22),th["font_body"],13,rgb(th["body_color"]),wrap=True)
        se.append((pt_,"fly",d+100))
    inject(slide,se); morph(slide); notes(slide,data.get("notes","")); return slide

def layout_numbered(prs, data, th, idx):
    """Left panel + right numbered points with motion path."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, rgb(th["bg"])); se=[]
    panel=rect(slide,Inches(0),Inches(0),Inches(5.6),H,rgb(th["header_bg"]))
    se.append((panel,"fade",0))
    sep=rect(slide,Inches(5.6),Inches(0),Inches(0.07),H,rgb(th["accent"]))
    se.append((sep,"wipe",100))
    n=txt(slide,f"{idx:02d}",Inches(0.4),Inches(0.3),Inches(1),Inches(0.4),th["font_body"],10,rgb(th["accent"]))
    se.append((n,"appear",150))
    t=txt(slide,data["title"],Inches(0.4),Inches(0.8),Inches(4.9),Inches(1.5),th["font_head"],32,rgb(th["title_color"]),bold=True)
    se.append((t,"fly",240))
    s=txt(slide,data.get("subtitle",""),Inches(0.4),Inches(2.35),Inches(4.9),Inches(0.6),th["font_body"],13,rgb(th["accent"]),italic=True)
    se.append((s,"fade",390))
    d=txt(slide,data.get("detail",""),Inches(0.4),Inches(3.1),Inches(4.9),Inches(4.0),th["font_body"],13,rgb(th["body_color"]),wrap=True)
    se.append((d,"fade",490))
    for i,pt in enumerate(data.get("points",[])[:4]):
        y=Inches(0.55+i*1.65); dd=300+i*160
        circ=rect(slide,Inches(6.05),y,Inches(0.72),Inches(0.72),rgb(th["accent"]))
        se.append((circ,"zoom",dd))
        nt=txt(slide,str(i+1),Inches(6.05),y,Inches(0.72),Inches(0.72),th["font_head"],18,rgb(th["bg"]),bold=True,align=PP_ALIGN.CENTER)
        se.append((nt,"appear",dd+20))
        pt_=txt(slide,pt,Inches(6.95),y,Inches(6.0),Inches(1.5),th["font_body"],13,rgb(th["body_color"]),wrap=True)
        se.append((pt_,"path",dd+80))
    inject(slide,se); morph(slide); notes(slide,data.get("notes","")); return slide

def layout_timeline(prs, data, th, idx):
    """Horizontal timeline with Wipe line and Zoom dots."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, rgb(th["bg"])); se=[]
    n=txt(slide,f"{idx:02d}",Inches(0.4),Inches(0.2),Inches(1),Inches(0.4),th["font_body"],10,rgb(th["accent"]))
    se.append((n,"appear",0))
    t=txt(slide,data["title"],Inches(0.4),Inches(0.42),Inches(12.5),Inches(1.1),th["font_head"],38,rgb(th["title_color"]),bold=True)
    se.append((t,"fly",100))
    s=txt(slide,data.get("subtitle",""),Inches(0.4),Inches(1.52),Inches(10),Inches(0.47),th["font_body"],13,rgb(th["accent"]),italic=True)
    se.append((s,"fade",280))
    line=rect(slide,Inches(0.4),Inches(3.56),Inches(12.5),Inches(0.07),rgb(th["accent"]))
    se.append((line,"wipe",380))
    pts=data.get("points",[]); n_pts=min(len(pts),4)
    sp=Inches(12.4)/max(n_pts,1)
    for i,pt in enumerate(pts[:4]):
        cx=Inches(0.4)+i*sp+sp/2-Inches(1.52); dd=500+i*160
        dot=rect(slide,cx+Inches(1.15),Inches(3.26),Inches(0.63),Inches(0.63),rgb(th["accent"]))
        se.append((dot,"zoom",dd))
        sn=txt(slide,str(i+1),cx+Inches(1.15),Inches(3.26),Inches(0.63),Inches(0.63),th["font_head"],14,rgb(th["bg"]),bold=True,align=PP_ALIGN.CENTER)
        se.append((sn,"appear",dd+20))
        card=rect(slide,cx,Inches(4.06),Inches(3.06),Inches(2.96),rgb(th["card_bg"]))
        se.append((card,"fade",dd+60))
        pt_=txt(slide,pt,cx+Inches(0.15),Inches(4.2),Inches(2.76),Inches(2.65),th["font_body"],12,rgb(th["body_color"]),wrap=True)
        se.append((pt_,"fly",dd+120))
    inject(slide,se); morph(slide); notes(slide,data.get("notes","")); return slide

def layout_spotlight(prs, data, th, idx):
    """Two-column bullets + bottom insight box with Spin accent."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, rgb(th["bg"])); se=[]
    bar=rect(slide,Inches(0),Inches(0),Inches(0.45),H,rgb(th["accent"]))
    se.append((bar,"wipe",0))
    n=txt(slide,f"{idx:02d}",Inches(0.65),Inches(0.3),Inches(1),Inches(0.45),th["font_body"],10,rgb(th["accent"]))
    se.append((n,"appear",100))
    t=txt(slide,data["title"],Inches(0.65),Inches(0.55),Inches(11.5),Inches(1.2),th["font_head"],38,rgb(th["title_color"]),bold=True)
    se.append((t,"float",200))
    s=txt(slide,data.get("subtitle",""),Inches(0.65),Inches(1.72),Inches(10),Inches(0.5),th["font_body"],14,rgb(th["accent"]),italic=True)
    se.append((s,"fade",350))
    for i,pt in enumerate(data.get("points",[])[:2]):
        b=txt(slide,"▸  "+pt,Inches(0.65),Inches(2.42+i*1.5),Inches(6.0),Inches(1.4),th["font_body"],13,rgb(th["body_color"]),wrap=True)
        se.append((b,"fly",450+i*130))
    for i,pt in enumerate(data.get("points",[])[2:4]):
        b=txt(slide,"▸  "+pt,Inches(7.0),Inches(2.42+i*1.5),Inches(6.0),Inches(1.4),th["font_body"],13,rgb(th["body_color"]),wrap=True)
        se.append((b,"fly",450+i*130))
    dbox=rect(slide,Inches(0.65),Inches(5.52),Inches(12.05),Inches(1.65),rgb(th["card_bg"]))
    se.append((dbox,"fade",700))
    dot=rect(slide,Inches(0.65),Inches(5.55),Inches(0.12),Inches(1.58),rgb(th["accent"]))
    se.append((dot,"spin",760))
    lbl=txt(slide,"KEY INSIGHT",Inches(0.9),Inches(5.63),Inches(3),Inches(0.35),th["font_body"],9,rgb(th["accent"]),bold=True)
    se.append((lbl,"appear",790))
    d=txt(slide,data.get("detail",""),Inches(0.9),Inches(5.98),Inches(11.8),Inches(1.05),th["font_body"],13,rgb(th["body_color"]),wrap=True)
    se.append((d,"fade",840))
    inject(slide,se); morph(slide); notes(slide,data.get("notes","")); return slide

# Layout sequences — varied per slide position
LAYOUTS = [layout_hero, layout_split, layout_grid, layout_numbered, layout_timeline, layout_spotlight]
LAYOUT_MAP = {
    "title_hero":   layout_hero,
    "two_column":   layout_split,
    "icon_grid":    layout_grid,
    "stat_callout": layout_numbered,
    "timeline":     layout_timeline,
    "full_detail":  layout_spotlight,
}

# ── Main builder ──────────────────────────────────────────────────────────────

def build_ppt(outline: list, theme: dict, style: str = "futuristic") -> bytes:
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    n = len(outline)
    for idx, data in enumerate(outline, start=1):
        reset_ids()
        # First and last slides always use hero
        if idx == 1 or idx == n:
            builder = layout_hero
        else:
            # Pick layout from AI suggestion or cycle through all 6
            key = data.get("layout", "")
            builder = LAYOUT_MAP.get(key)
            if not builder:
                # Rotate through remaining layouts (never repeat consecutively)
                builder = LAYOUTS[((idx-1) % (len(LAYOUTS)-1)) + 1]
        builder(prs, data, theme, idx)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.read()