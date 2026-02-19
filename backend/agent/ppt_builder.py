"""
ppt_builder.py  –  Antigravity PPT Engine
==========================================
Animations implemented (all via real OOXML):

  1. Morph Transition    – slide-level transition (morphTransition XML)
  2. Fade / Appear       – entrance: presetID 10 (fade), presetID 1 (appear)
  3. Fly In / Float In   – entrance: presetID 21 sub 4 (fly-from-bottom),
                                     presetID 15 sub 0 (float up)
  4. Zoom                – entrance: presetID 36 (zoom)
  5. Grow / Shrink       – emphasis: presetID 150 (grow/shrink scale)
  6. Custom Motion Path  – motion path along a straight line
  7. Spin / Teeter       – emphasis: presetID 155 (spin), presetID 156 (teeter)
  8. Wipe                – entrance: presetID 27 sub 8 (wipe from left)
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree
import io

# ── colour helper ─────────────────────────────────────────────────────────────
def rgb(h: str) -> RGBColor:
    h = h.lstrip("#")
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

# ── slide helpers ─────────────────────────────────────────────────────────────
def set_bg(slide, color: RGBColor):
    f = slide.background.fill
    f.solid(); f.fore_color.rgb = color

def add_rect(slide, l, t, w, h, fill_color: RGBColor):
    s = slide.shapes.add_shape(1, l, t, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = fill_color
    s.line.fill.background()
    return s

def add_text(slide, text, l, t, w, h,
             fn="Calibri", sz=18, color=RGBColor(255,255,255),
             bold=False, italic=False, align=PP_ALIGN.LEFT, wrap=True):
    txb = slide.shapes.add_textbox(l, t, w, h)
    tf  = txb.text_frame; tf.word_wrap = wrap
    p   = tf.paragraphs[0]; p.alignment = align
    r   = p.add_run(); r.text = text
    r.font.name = fn; r.font.size = Pt(sz)
    r.font.color.rgb = color; r.font.bold = bold; r.font.italic = italic
    return txb

def add_bullet_para(tf, text, fn, sz, color, space_before=10):
    p = tf.add_paragraph(); p.alignment = PP_ALIGN.LEFT
    p.space_before = Pt(space_before)
    r = p.add_run(); r.text = text
    r.font.name = fn; r.font.size = Pt(sz); r.font.color.rgb = color
    return p

def _add_notes(slide, txt):
    if txt:
        try: slide.notes_slide.notes_text_frame.text = txt
        except: pass

# ── global ID counter (reset per slide) ──────────────────────────────────────
_id = [10]
def nid(n=1):
    ids = list(range(_id[0], _id[0]+n))
    _id[0] += n
    return ids[0] if n == 1 else ids

# ════════════════════════════════════════════════════════════════════════════
# ANIMATION XML BUILDERS
# Each returns an XML string for one animation action.
# All are "withPrev" so they fire together on ONE click,
# staggered by the `delay` (ms) parameter.
# ════════════════════════════════════════════════════════════════════════════

P = "http://schemas.openxmlformats.org/presentationml/2006/main"

def _entrance_xml(spid, preset_id, preset_sub, dur, delay, grp, filt=""):
    i1,i2,i3,i4 = nid(4)
    anim_eff = ""
    if filt:
        anim_eff = f"""<p:animEffect transition="in" filter="{filt}">
          <p:cBhvr><p:cTn id="{i4}" dur="{dur}"/>
            <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
          </p:cBhvr></p:animEffect>"""
    return f"""
<p:par><p:cTn id="{i1}" presetID="{preset_id}" presetClass="entr"
  presetSubtype="{preset_sub}" fill="hold" grpId="{grp}" nodeType="withEffect">
  <p:stCondLst><p:cond delay="{delay}"/></p:stCondLst>
  <p:childTnLst>
    <p:set><p:cBhvr>
        <p:cTn id="{i2}" dur="1" fill="hold"/>
        <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
        <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
      </p:cBhvr><p:to><p:strVal val="visible"/></p:to>
    </p:set>
    {anim_eff}
  </p:childTnLst>
</p:cTn></p:par>"""

def anim_fade(spid, delay, grp):
    """Fade entrance – presetID 10"""
    return _entrance_xml(spid, 10, 0, 600, delay, grp, filt="fade")

def anim_appear(spid, delay, grp):
    """Appear entrance – presetID 1"""
    return _entrance_xml(spid, 1, 0, 1, delay, grp)

def anim_fly_in(spid, delay, grp):
    """Fly In from Bottom – presetID 21 sub 4"""
    return _entrance_xml(spid, 21, 4, 600, delay, grp, filt="fade")

def anim_float_in(spid, delay, grp):
    """Float Up – presetID 15"""
    return _entrance_xml(spid, 15, 0, 700, delay, grp, filt="fade")

def anim_zoom(spid, delay, grp):
    """Zoom entrance – presetID 36"""
    return _entrance_xml(spid, 36, 0, 500, delay, grp, filt="fade")

def anim_wipe(spid, delay, grp):
    """Wipe from Left – presetID 27 sub 8"""
    return _entrance_xml(spid, 27, 8, 500, delay, grp, filt="wipe(right)")

def anim_grow_shrink(spid, delay, grp):
    """Grow/Shrink emphasis – animates scaleX/Y to 115% then back"""
    i1,i2,i3,i4,i5,i6 = nid(6)
    return f"""
<p:par><p:cTn id="{i1}" presetID="150" presetClass="emph"
  presetSubtype="0" fill="hold" grpId="{grp}" nodeType="withEffect">
  <p:stCondLst><p:cond delay="{delay}"/></p:stCondLst>
  <p:childTnLst>
    <p:animScale>
      <p:cBhvr><p:cTn id="{i2}" dur="400" autoRev="1"/>
        <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
      </p:cBhvr>
      <p:by x="115000" y="115000"/>
    </p:animScale>
  </p:childTnLst>
</p:cTn></p:par>"""

def anim_spin(spid, delay, grp):
    """Spin emphasis – 360-degree clockwise rotation"""
    i1,i2 = nid(2)
    return f"""
<p:par><p:cTn id="{i1}" presetID="155" presetClass="emph"
  presetSubtype="0" fill="hold" grpId="{grp}" nodeType="withEffect">
  <p:stCondLst><p:cond delay="{delay}"/></p:stCondLst>
  <p:childTnLst>
    <p:animRot><p:cBhvr>
        <p:cTn id="{i2}" dur="600"/>
        <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
      </p:cBhvr>
      <p:by ang="21600000"/>
    </p:animRot>
  </p:childTnLst>
</p:cTn></p:par>"""

def anim_teeter(spid, delay, grp):
    """Teeter (rock side-to-side) emphasis"""
    i1,i2,i3 = nid(3)
    return f"""
<p:par><p:cTn id="{i1}" presetID="156" presetClass="emph"
  presetSubtype="0" fill="hold" grpId="{grp}" nodeType="withEffect">
  <p:stCondLst><p:cond delay="{delay}"/></p:stCondLst>
  <p:childTnLst>
    <p:animRot><p:cBhvr>
        <p:cTn id="{i2}" dur="200" autoRev="1" repeatCount="2"/>
        <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
      </p:cBhvr>
      <p:by ang="1080000"/>
    </p:animRot>
  </p:childTnLst>
</p:cTn></p:par>"""

def anim_motion_path(spid, delay, grp, dx_emu=914400, dy_emu=0):
    """
    Custom motion path – moves shape by (dx_emu, dy_emu) from its position.
    Default: slide in 1 inch from the left (negative x start).
    Path coords are fractions of slide size (0..1 = 0%..100%).
    """
    i1,i2 = nid(2)
    # slide W=9144000 EMU, H=5143500 EMU (13.33" x 7.5")
    SW = 9144000; SH = 5143500
    sx = -dx_emu / SW;  sy = dy_emu / SH
    ex = 0.0;           ey = 0.0
    path_str = f"M {sx:.4f} {sy:.4f} L {ex:.4f} {ey:.4f}"
    return f"""
<p:par><p:cTn id="{i1}" presetID="0" presetClass="entr"
  presetSubtype="0" fill="hold" grpId="{grp}" nodeType="withEffect">
  <p:stCondLst><p:cond delay="{delay}"/></p:stCondLst>
  <p:childTnLst>
    <p:animMotion origin="layout" path="{path_str}" pathEditMode="auto">
      <p:cBhvr><p:cTn id="{i2}" dur="700" fill="hold"/>
        <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
        <p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst>
      </p:cBhvr>
    </p:animMotion>
  </p:childTnLst>
</p:cTn></p:par>"""

# ════════════════════════════════════════════════════════════════════════════
# MORPH TRANSITION  (slide-level – not shape animation)
# ════════════════════════════════════════════════════════════════════════════

def apply_morph_transition(slide):
    """
    Adds a Morph transition to the slide.
    This is the <p:transition> element with morphTransition child.
    """
    try:
        slide_el = slide._element
        # Remove existing transition if any
        existing = slide_el.find(qn("p:transition"))
        if existing is not None:
            slide_el.remove(existing)

        trans_xml = """<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                                     xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"
                                     spd="med" advTm="0">
  <p14:morphTransition xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"
                        option="byObject"/>
</p:transition>"""
        trans_el = etree.fromstring(trans_xml.encode("utf-8"))
        # Insert transition before timing element if exists, else append
        timing = slide_el.find(qn("p:timing"))
        if timing is not None:
            timing.addprevious(trans_el)
        else:
            slide_el.append(trans_el)
    except Exception as e:
        print(f"[Morph Warning] {e}")

# ════════════════════════════════════════════════════════════════════════════
# TIMING TREE INJECTOR
# ════════════════════════════════════════════════════════════════════════════

def inject_animations(slide, anim_blocks: list):
    """
    anim_blocks: list of XML strings (one per shape animation).
    All wrapped in a single onClick trigger → all fire on one click.
    """
    if not anim_blocks:
        return
    inner = "\n".join(anim_blocks)
    i1,i2,i3 = nid(3)
    timing_xml = f"""<p:timing
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:tnLst>
    <p:par>
      <p:cTn id="{i1}" dur="indefinite" restart="whenNotActive" nodeType="tmRoot">
        <p:childTnLst>
          <p:par>
            <p:cTn id="{i2}" fill="hold">
              <p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>
              <p:childTnLst>
                <p:par>
                  <p:cTn id="{i3}" fill="hold" nodeType="clickEffect">
                    <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                    <p:childTnLst>
                      {inner}
                    </p:childTnLst>
                  </p:cTn>
                </p:par>
              </p:childTnLst>
            </p:cTn>
          </p:par>
        </p:childTnLst>
      </p:cTn>
    </p:par>
  </p:tnLst>
  <p:bldLst/>
</p:timing>"""
    try:
        slide_el = slide._element
        ex = slide_el.find(qn("p:timing"))
        if ex is not None: slide_el.remove(ex)
        slide_el.append(etree.fromstring(timing_xml.encode("utf-8")))
    except Exception as e:
        print(f"[Timing Warning] {e}")

# ════════════════════════════════════════════════════════════════════════════
# SLIDE DIMENSIONS
# ════════════════════════════════════════════════════════════════════════════
W = Inches(13.33)
H = Inches(7.5)

# ════════════════════════════════════════════════════════════════════════════
# SLIDE LAYOUT BUILDERS
# Each layout uses a different mix of the 7 animation types
# ════════════════════════════════════════════════════════════════════════════

def slide_title_hero(prs, data, theme, idx):
    """
    Hero title slide.
    Animations: Morph transition + Wipe (bar) + Zoom (title) +
                Fade (subtitle) + Float In (detail) + Spin (badge circle)
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, rgb(theme["bg"]))
    _id[0] = 10; ab = []; g = 1

    bar = add_rect(slide, Inches(0), Inches(0), Inches(0.45), H, rgb(theme["accent"]))
    ab.append(anim_wipe(bar.shape_id, 0, g)); g+=1

    circ = slide.shapes.add_shape(9, Inches(9.3), Inches(-1.8), Inches(5.5), Inches(5.5))
    circ.fill.solid(); circ.fill.fore_color.rgb = rgb(theme["accent"])
    circ.line.fill.background()
    ab.append(anim_fade(circ.shape_id, 0, g)); g+=1
    ab.append(anim_spin(circ.shape_id, 200, g)); g+=1   # SPIN on the decorative circle

    num = add_text(slide, f"{idx:02d}", Inches(0.7), Inches(0.3), Inches(1.2), Inches(0.5),
                   theme["font_body"], 11, rgb(theme["accent"]))
    ab.append(anim_appear(num.shape_id, 150, g)); g+=1

    title = add_text(slide, data["title"], Inches(0.7), Inches(1.3), Inches(10.5), Inches(2.7),
                     theme["font_head"], 54, rgb(theme["title_color"]), bold=True)
    ab.append(anim_zoom(title.shape_id, 250, g)); g+=1  # ZOOM on main title

    sub = add_text(slide, data.get("subtitle",""), Inches(0.7), Inches(3.85), Inches(10), Inches(0.8),
                   theme["font_body"], 20, rgb(theme["accent"]), italic=True)
    ab.append(anim_fade(sub.shape_id, 500, g)); g+=1    # FADE subtitle

    detail = add_text(slide, data.get("detail",""), Inches(0.7), Inches(4.75), Inches(10.5), Inches(2.0),
                      theme["font_body"], 14, rgb(theme["muted_color"]), wrap=True)
    ab.append(anim_float_in(detail.shape_id, 700, g)); g+=1  # FLOAT IN detail

    inject_animations(slide, ab)
    apply_morph_transition(slide)                        # MORPH transition
    _add_notes(slide, data.get("notes",""))
    return slide


def slide_two_column(prs, data, theme, idx):
    """
    Two-column layout.
    Animations: Morph + Wipe (top bar) + Fly In (title) +
                Motion Path (bullets slide in) + Fade (right card)
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, rgb(theme["bg"]))
    _id[0] = 10; ab = []; g = 1

    bar = add_rect(slide, Inches(0), Inches(0), W, Inches(0.07), rgb(theme["accent"]))
    ab.append(anim_wipe(bar.shape_id, 0, g)); g+=1

    num = add_text(slide, f"{idx:02d}", Inches(0.4), Inches(0.22), Inches(1), Inches(0.4),
                   theme["font_body"], 10, rgb(theme["accent"]))
    ab.append(anim_appear(num.shape_id, 80, g)); g+=1

    title = add_text(slide, data["title"], Inches(0.4), Inches(0.5), Inches(7.5), Inches(1.2),
                     theme["font_head"], 36, rgb(theme["title_color"]), bold=True)
    ab.append(anim_fly_in(title.shape_id, 180, g)); g+=1   # FLY IN title

    sub = add_text(slide, data.get("subtitle",""), Inches(0.4), Inches(1.65), Inches(7), Inches(0.5),
                   theme["font_body"], 13, rgb(theme["accent"]), italic=True)
    ab.append(anim_fade(sub.shape_id, 360, g)); g+=1

    # Bullets – MOTION PATH (slide in from left)
    btxb = slide.shapes.add_textbox(Inches(0.4), Inches(2.3), Inches(6.5), Inches(4.8))
    btf  = btxb.text_frame; btf.word_wrap = True
    for i, pt in enumerate(data.get("points",[])):
        if i == 0:
            p0 = btf.paragraphs[0]; p0.alignment = PP_ALIGN.LEFT
            r = p0.add_run(); r.text = "▸  " + pt
            r.font.name = theme["font_body"]; r.font.size = Pt(13)
            r.font.color.rgb = rgb(theme["body_color"])
        else:
            add_bullet_para(btf, "▸  " + pt, theme["font_body"], 13, rgb(theme["body_color"]))
    ab.append(anim_motion_path(btxb.shape_id, 460, g,
                               dx_emu=int(Inches(2)), dy_emu=0)); g+=1  # MOTION PATH bullets

    # Right card – FADE
    card = add_rect(slide, Inches(7.3), Inches(0.4), Inches(5.65), Inches(6.7), rgb(theme["card_bg"]))
    ab.append(anim_fade(card.shape_id, 260, g)); g+=1

    cbar = add_rect(slide, Inches(7.3), Inches(0.4), Inches(0.1), Inches(6.7), rgb(theme["accent"]))
    ab.append(anim_wipe(cbar.shape_id, 300, g)); g+=1

    lbl = add_text(slide, "DEEP DIVE", Inches(7.6), Inches(0.78), Inches(5), Inches(0.38),
                   theme["font_body"], 9, rgb(theme["accent"]), bold=True)
    ab.append(anim_appear(lbl.shape_id, 380, g)); g+=1

    detail = add_text(slide, data.get("detail",""), Inches(7.6), Inches(1.28), Inches(5.1), Inches(5.65),
                      theme["font_body"], 14, rgb(theme["body_color"]), wrap=True)
    ab.append(anim_float_in(detail.shape_id, 520, g)); g+=1   # FLOAT IN deep-dive text

    inject_animations(slide, ab)
    apply_morph_transition(slide)
    _add_notes(slide, data.get("notes",""))
    return slide


def slide_icon_grid(prs, data, theme, idx):
    """
    2×2 card grid.
    Animations: Morph + Fly In (header) + Zoom (each card) +
                Grow/Shrink (badge numbers) + Wipe (reveal point text)
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, rgb(theme["bg"]))
    _id[0] = 10; ab = []; g = 1

    strip = add_rect(slide, Inches(0), Inches(0), W, Inches(2.1), rgb(theme["header_bg"]))
    ab.append(anim_fade(strip.shape_id, 0, g)); g+=1

    num = add_text(slide, f"{idx:02d}", Inches(0.4), Inches(0.2), Inches(1), Inches(0.4),
                   theme["font_body"], 10, rgb(theme["accent"]))
    ab.append(anim_appear(num.shape_id, 60, g)); g+=1

    title = add_text(slide, data["title"], Inches(0.4), Inches(0.42), Inches(12.5), Inches(1.1),
                     theme["font_head"], 38, rgb(theme["title_color"]), bold=True)
    ab.append(anim_fly_in(title.shape_id, 150, g)); g+=1      # FLY IN

    sub = add_text(slide, data.get("subtitle",""), Inches(0.4), Inches(1.52), Inches(10), Inches(0.47),
                   theme["font_body"], 13, rgb(theme["accent"]), italic=True)
    ab.append(anim_fade(sub.shape_id, 320, g)); g+=1

    points   = data.get("points", [""] * 4)
    positions = [
        (Inches(0.3),  Inches(2.25)),
        (Inches(6.85), Inches(2.25)),
        (Inches(0.3),  Inches(4.95)),
        (Inches(6.85), Inches(4.95)),
    ]
    cw, ch = Inches(6.25), Inches(2.45)

    for i, (pt, (cx, cy)) in enumerate(zip(points[:4], positions)):
        d = 420 + i * 130

        card = add_rect(slide, cx, cy, cw, ch, rgb(theme["card_bg"]))
        ab.append(anim_zoom(card.shape_id, d, g)); g+=1        # ZOOM each card

        badge = add_rect(slide, cx+Inches(0.15), cy+Inches(0.13),
                         Inches(0.56), Inches(0.56), rgb(theme["accent"]))
        ab.append(anim_grow_shrink(badge.shape_id, d+60, g)); g+=1  # GROW/SHRINK badge

        bnum = add_text(slide, f"0{i+1}", cx+Inches(0.15), cy+Inches(0.13),
                        Inches(0.56), Inches(0.56), theme["font_body"], 12,
                        rgb(theme["bg"]), bold=True, align=PP_ALIGN.CENTER)
        ab.append(anim_appear(bnum.shape_id, d+60, g)); g+=1

        ptxt = add_text(slide, pt, cx+Inches(0.85), cy+Inches(0.1),
                        cw-Inches(1.05), ch-Inches(0.22),
                        theme["font_body"], 13, rgb(theme["body_color"]), wrap=True)
        ab.append(anim_wipe(ptxt.shape_id, d+100, g)); g+=1   # WIPE reveal text

    inject_animations(slide, ab)
    apply_morph_transition(slide)
    _add_notes(slide, data.get("notes",""))
    return slide


def slide_stat_callout(prs, data, theme, idx):
    """
    Left panel + numbered right list.
    Animations: Morph + Fade (panel) + Zoom (title) +
                Teeter (number circles) + Fly In (point text)
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, rgb(theme["bg"]))
    _id[0] = 10; ab = []; g = 1

    panel = add_rect(slide, Inches(0), Inches(0), Inches(5.6), H, rgb(theme["header_bg"]))
    ab.append(anim_fade(panel.shape_id, 0, g)); g+=1

    sep = add_rect(slide, Inches(5.6), Inches(0), Inches(0.07), H, rgb(theme["accent"]))
    ab.append(anim_wipe(sep.shape_id, 80, g)); g+=1

    num_lbl = add_text(slide, f"{idx:02d}", Inches(0.4), Inches(0.3), Inches(1), Inches(0.4),
                       theme["font_body"], 10, rgb(theme["accent"]))
    ab.append(anim_appear(num_lbl.shape_id, 120, g)); g+=1

    title = add_text(slide, data["title"], Inches(0.4), Inches(0.8), Inches(4.9), Inches(1.5),
                     theme["font_head"], 32, rgb(theme["title_color"]), bold=True)
    ab.append(anim_zoom(title.shape_id, 220, g)); g+=1         # ZOOM title

    sub = add_text(slide, data.get("subtitle",""), Inches(0.4), Inches(2.38), Inches(4.9), Inches(0.6),
                   theme["font_body"], 13, rgb(theme["accent"]), italic=True)
    ab.append(anim_fade(sub.shape_id, 380, g)); g+=1

    detail = add_text(slide, data.get("detail",""), Inches(0.4), Inches(3.1), Inches(4.9), Inches(4.0),
                      theme["font_body"], 13, rgb(theme["body_color"]), wrap=True)
    ab.append(anim_float_in(detail.shape_id, 480, g)); g+=1    # FLOAT IN detail

    for i, pt in enumerate(data.get("points",[])[:4]):
        y  = Inches(0.55 + i * 1.65)
        d  = 280 + i * 150

        circ = add_rect(slide, Inches(6.0), y, Inches(0.72), Inches(0.72), rgb(theme["accent"]))
        ab.append(anim_appear(circ.shape_id, d, g)); g+=1
        ab.append(anim_teeter(circ.shape_id, d+100, g)); g+=1  # TEETER number circle

        ntxt = add_text(slide, str(i+1), Inches(6.0), y, Inches(0.72), Inches(0.72),
                        theme["font_head"], 18, rgb(theme["bg"]), bold=True, align=PP_ALIGN.CENTER)
        ab.append(anim_appear(ntxt.shape_id, d, g)); g+=1

        ptxt = add_text(slide, pt, Inches(6.9), y, Inches(6.05), Inches(1.5),
                        theme["font_body"], 13, rgb(theme["body_color"]), wrap=True)
        ab.append(anim_fly_in(ptxt.shape_id, d+80, g)); g+=1   # FLY IN each point

    inject_animations(slide, ab)
    apply_morph_transition(slide)
    _add_notes(slide, data.get("notes",""))
    return slide


def slide_timeline(prs, data, theme, idx):
    """
    Horizontal timeline.
    Animations: Morph + Fly In (title) + Wipe (timeline line) +
                Motion Path (cards slide up) + Grow/Shrink (dots)
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, rgb(theme["bg"]))
    _id[0] = 10; ab = []; g = 1

    num = add_text(slide, f"{idx:02d}", Inches(0.4), Inches(0.2), Inches(1), Inches(0.4),
                   theme["font_body"], 10, rgb(theme["accent"]))
    ab.append(anim_appear(num.shape_id, 0, g)); g+=1

    title = add_text(slide, data["title"], Inches(0.4), Inches(0.42), Inches(12.5), Inches(1.1),
                     theme["font_head"], 38, rgb(theme["title_color"]), bold=True)
    ab.append(anim_fly_in(title.shape_id, 80, g)); g+=1        # FLY IN title

    sub = add_text(slide, data.get("subtitle",""), Inches(0.4), Inches(1.52), Inches(10), Inches(0.47),
                   theme["font_body"], 13, rgb(theme["accent"]), italic=True)
    ab.append(anim_fade(sub.shape_id, 260, g)); g+=1

    line = add_rect(slide, Inches(0.4), Inches(3.56), Inches(12.5), Inches(0.07), rgb(theme["accent"]))
    ab.append(anim_wipe(line.shape_id, 350, g)); g+=1          # WIPE timeline line

    points = data.get("points", [])
    n = min(len(points), 4)
    spacing = Inches(12.4) / max(n, 1)

    for i, pt in enumerate(points[:4]):
        cx = Inches(0.4) + i * spacing + spacing/2 - Inches(1.52)
        d  = 450 + i * 150

        dot = add_rect(slide, cx+Inches(1.15), Inches(3.26), Inches(0.63), Inches(0.63), rgb(theme["accent"]))
        ab.append(anim_appear(dot.shape_id, d, g)); g+=1
        ab.append(anim_grow_shrink(dot.shape_id, d+80, g)); g+=1  # GROW/SHRINK dot

        snum = add_text(slide, str(i+1), cx+Inches(1.15), Inches(3.26), Inches(0.63), Inches(0.63),
                        theme["font_head"], 14, rgb(theme["bg"]), bold=True, align=PP_ALIGN.CENTER)
        ab.append(anim_appear(snum.shape_id, d, g)); g+=1

        card = add_rect(slide, cx, Inches(4.06), Inches(3.06), Inches(2.96), rgb(theme["card_bg"]))
        # MOTION PATH – cards slide up from below
        ab.append(anim_motion_path(card.shape_id, d+60, g,
                                   dx_emu=0, dy_emu=int(Inches(1.5)))); g+=1

        ptxt = add_text(slide, pt, cx+Inches(0.15), Inches(4.2), Inches(2.76), Inches(2.65),
                        theme["font_body"], 12, rgb(theme["body_color"]), wrap=True)
        ab.append(anim_fade(ptxt.shape_id, d+120, g)); g+=1

    inject_animations(slide, ab)
    apply_morph_transition(slide)
    _add_notes(slide, data.get("notes",""))
    return slide


def slide_full_detail(prs, data, theme, idx):
    """
    2-column bullets + insight box.
    Animations: Morph + Wipe (bar) + Float In (title) +
                Motion Path (left bullets) + Fly In (right bullets) +
                Zoom (insight box)
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, rgb(theme["bg"]))
    _id[0] = 10; ab = []; g = 1

    bar = add_rect(slide, Inches(0), Inches(0), Inches(0.45), H, rgb(theme["accent"]))
    ab.append(anim_wipe(bar.shape_id, 0, g)); g+=1

    num = add_text(slide, f"{idx:02d}", Inches(0.65), Inches(0.3), Inches(1), Inches(0.45),
                   theme["font_body"], 10, rgb(theme["accent"]))
    ab.append(anim_appear(num.shape_id, 80, g)); g+=1

    title = add_text(slide, data["title"], Inches(0.65), Inches(0.55), Inches(11.5), Inches(1.2),
                     theme["font_head"], 38, rgb(theme["title_color"]), bold=True)
    ab.append(anim_float_in(title.shape_id, 180, g)); g+=1     # FLOAT IN title

    sub = add_text(slide, data.get("subtitle",""), Inches(0.65), Inches(1.72), Inches(10), Inches(0.5),
                   theme["font_body"], 14, rgb(theme["accent"]), italic=True)
    ab.append(anim_fade(sub.shape_id, 340, g)); g+=1

    # Left bullets – MOTION PATH from left
    for i, pt in enumerate(data.get("points", [])[:2]):
        b = add_text(slide, "▸  " + pt, Inches(0.65), Inches(2.42 + i*1.5), Inches(6.0), Inches(1.4),
                     theme["font_body"], 13, rgb(theme["body_color"]), wrap=True)
        ab.append(anim_motion_path(b.shape_id, 440 + i*130, g,
                                   dx_emu=int(Inches(2)), dy_emu=0)); g+=1

    # Right bullets – FLY IN
    for i, pt in enumerate(data.get("points", [])[2:4]):
        b = add_text(slide, "▸  " + pt, Inches(7.0), Inches(2.42 + i*1.5), Inches(6.0), Inches(1.4),
                     theme["font_body"], 13, rgb(theme["body_color"]), wrap=True)
        ab.append(anim_fly_in(b.shape_id, 440 + i*130, g)); g+=1

    # Insight box – ZOOM
    dbox = add_rect(slide, Inches(0.65), Inches(5.52), Inches(12.05), Inches(1.62), rgb(theme["card_bg"]))
    ab.append(anim_zoom(dbox.shape_id, 680, g)); g+=1

    dlbl = add_text(slide, "KEY INSIGHT", Inches(0.9), Inches(5.63), Inches(3), Inches(0.35),
                    theme["font_body"], 9, rgb(theme["accent"]), bold=True)
    ab.append(anim_appear(dlbl.shape_id, 730, g)); g+=1

    dtxt = add_text(slide, data.get("detail",""), Inches(0.9), Inches(5.98), Inches(11.8), Inches(1.0),
                    theme["font_body"], 13, rgb(theme["body_color"]), wrap=True)
    ab.append(anim_fade(dtxt.shape_id, 780, g)); g+=1

    inject_animations(slide, ab)
    apply_morph_transition(slide)
    _add_notes(slide, data.get("notes",""))
    return slide


# ════════════════════════════════════════════════════════════════════════════
# ROUTER + MAIN ENTRY POINT
# ════════════════════════════════════════════════════════════════════════════

LAYOUT_BUILDERS = {
    "title_hero":   slide_title_hero,
    "two_column":   slide_two_column,
    "icon_grid":    slide_icon_grid,
    "stat_callout": slide_stat_callout,
    "timeline":     slide_timeline,
    "full_detail":  slide_full_detail,
}

def build_ppt(outline: list, theme: dict, style: str = "futuristic") -> bytes:
    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)

    for idx, slide_data in enumerate(outline, start=1):
        _id[0] = 10  # reset per slide
        layout_key = slide_data.get("layout", "two_column")
        builder    = LAYOUT_BUILDERS.get(layout_key, slide_two_column)
        builder(prs, slide_data, theme, idx)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.read()