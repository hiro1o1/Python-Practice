"""Catalog-wide redesign: depth, lighting, section openers and readability.

    python catalog/tools/restyle.py <in.pptx> <out.pptx>

Nothing is deleted. The pass
* tints light pages so white cards and product stages stand out, with soft
  top-lit gradients and clear shadows
* grounds product cut-outs with a contact shadow and rounds photo corners
* adds a dark opener page before each product line (hero product, pages in
  the section, and a quote / OEM-ODM call to action from the contact page)
* adds glows behind hero products, line icons on stat cards, banded tables
* enlarges small text, uses the accent colour for port markers, and unifies
  "Type-C" as "USB-C"
* keeps page numbers and the page references on page 4 correct
"""

import copy
import io
import os
import sys

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.util import Mm, Pt

ICONS = os.path.join(os.path.dirname(__file__), "icons")

PAGE_BG = "F2F5F8"
NAVY, INK, BODY, MUTED, TEAL, ACCENT, SKY, PALE = ("0B1B2B", "0B1B2B", "1B2733", "566270",
                                                   "09678F", "0E9BD6", "6CC8EC", "B4C2D1")
STAGES = {"Product stage", "Card stage", "Figure tile"}
CARDS = {"Card", "Use case", "Mainboard card", "Fact tile", "Statement panel"}
WHITE_PANELS = {"Image panel", "Series stage", "Drawing frame"}
MIN_SIZE = {"Caption": 900, "Label": 900, "Platform label": 900, "Card note": 850, "Use case note": 850,
            "Fact label": 850, "Highlight label": 900, "Series finish": 850, "Spec label": 900,
            "Port legend": 900, "Page reference": 1000}
ICON_FOR = [("Since", "calendar"), ("50–100", "users"), ("RMB", "capital"), ("m²", "area"),
            ("80+", "globe"), ("4K", "monitor"), ("points", "touch"), ("ms", "zap"), ("°", "eye"),
            ("Core", "cpu"), ("TB", "storage"), ("GB", "memory"), ("Air-cooled", "fan"),
            ("Air conditioner", "snow"), ("MTK", "chip"), ("MP", "camera")]
TERMS = [("USB Type-C", "USB-C"), ("Type-C", "USB-C")]

# Product lines: first page of each (1-based, before openers are inserted) and its last page.
SECTIONS = [(5, 11), (12, 14), (15, 16), (17, 19)]
CTA_TITLE = "Request a quote or an OEM/ODM version"


# --------------------------------------------------------------------------- xml helpers
def el(tag, **attrs):
    e = etree.Element(qn(tag))
    for k, v in attrs.items():
        e.set(k, str(v))
    return e


def color(val, alpha=None):
    c = el("a:srgbClr", val=val)
    if alpha is not None:
        c.append(el("a:alpha", val=int(alpha * 1000)))
    return c


def solid(val):
    f = el("a:solidFill")
    f.append(color(val))
    return f


def shadow(blur_pt, dist_pt, alpha, val="0B1B2B"):
    lst = el("a:effectLst")
    s = el("a:outerShdw", blurRad=int(blur_pt * 12700), dist=int(dist_pt * 12700),
           dir=5400000, algn="t", rotWithShape=0)
    s.append(color(val, alpha))
    lst.append(s)
    return lst


def set_effect(spPr, lst):
    old = spPr.find(qn("a:effectLst"))
    if old is not None:
        spPr.replace(old, lst)
        return
    for tag in ("a:ln", "a:solidFill", "a:gradFill", "a:noFill", "a:blipFill", "a:prstGeom", "a:custGeom"):
        anchor = spPr.find(qn(tag))
        if anchor is not None:
            anchor.addnext(lst)
            return
    spPr.append(lst)


def set_fill(spPr, fill):
    for tag in ("a:solidFill", "a:gradFill", "a:noFill"):
        old = spPr.find(qn(tag))
        if old is not None:
            spPr.replace(old, fill)
            return
    geom = spPr.find(qn("a:prstGeom"))
    geom.addnext(fill)


def set_line(spPr, val, w=6350):
    ln = el("a:ln", w=w)
    ln.append(solid(val))
    old = spPr.find(qn("a:ln"))
    if old is not None:
        spPr.replace(old, ln)
    else:
        fill = spPr.find(qn("a:solidFill"))
        (fill if fill is not None else spPr.find(qn("a:gradFill"))).addnext(ln)


def gradient(top, bottom):
    g = el("a:gradFill", rotWithShape=1)
    gs = el("a:gsLst")
    for pos, c in ((0, top), (100000, bottom)):
        s = el("a:gs", pos=pos)
        s.append(color(c))
        gs.append(s)
    g.append(gs)
    g.append(el("a:lin", ang=5400000, scaled=0))
    return g


def run_color(r, val):
    f = r.find(qn("a:solidFill"))
    if f is not None:
        r.remove(f)
    r.insert(0, solid(val))


# --------------------------------------------------------------------------- shape helpers
def add_text(slide, x, y, w, h, text, size, val, bold=False, align="l", caps=False, spc=None, anchor="t"):
    tb = slide.shapes.add_textbox(Mm(x), Mm(y), Mm(w), Mm(h))
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.word_wrap = True
    bp = tf._txBody.find(qn("a:bodyPr"))
    bp.set("anchor", anchor)
    p = tf.paragraphs[0]
    p.alignment = {"l": 1, "c": 2, "r": 3}[align]
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.name = "Segoe UI Semibold" if bold else "Segoe UI"
    r.font.color.rgb = RGBColor.from_string(val)
    rpr = r._r.find(qn("a:rPr"))
    if caps:
        rpr.set("cap", "all")
    if spc:
        rpr.set("spc", str(spc))
    return tb


def add_box(slide, x, y, w, h, fill, radius=None, line=None):
    shape = slide.shapes.add_shape(5 if radius else 1, Mm(x), Mm(y), Mm(w), Mm(h))
    spPr = shape._element.spPr
    set_fill(spPr, solid(fill))
    if line:
        set_line(spPr, line)
    else:
        shape.line.fill.background()
    set_effect(spPr, el("a:effectLst"))
    if radius is not None:
        shape.adjustments[0] = radius
    return shape


def glow(slide, x, y, w, h, after=None, strength=1.0):
    pic = slide.shapes.add_picture(os.path.join(ICONS, "glow.png"), Mm(x), Mm(y), Mm(w), Mm(h))
    pic.name = "Glow"
    if strength < 1:
        blip = pic._element.find(".//" + qn("a:blip"))
        blip.append(el("a:alphaModFix", amt=int(strength * 100000)))
    if after is not None:
        after.addnext(pic._element)
    else:
        slide.shapes._spTree.insert(2, pic._element)
    return pic


def is_cutout(pic):
    try:
        im = Image.open(io.BytesIO(pic.image.blob))
    except Exception:
        return False
    if im.mode not in ("RGBA", "LA", "P"):
        return False
    a = im.convert("RGBA").getchannel("A")
    return a.getextrema()[0] < 200


def copy_picture(src_pic, slide, x, y, w, h):
    """Place an existing picture's image on another slide, fitted into a box."""
    iw, ih = src_pic.width, src_pic.height
    s = min(Mm(w) / iw, Mm(h) / ih)
    pw, ph = int(iw * s), int(ih * s)
    el_ = copy.deepcopy(src_pic._element)
    rid = slide.part.relate_to(src_pic.part.related_part(src_pic._element.blipFill.blip.rEmbed),
                               "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
    el_.blipFill.blip.set(qn("r:embed"), rid)
    off = el_.find(".//" + qn("a:off"))
    ext = el_.find(".//" + qn("a:ext"))
    off.set("x", str(int(Mm(x) + (Mm(w) - pw) / 2)))
    off.set("y", str(int(Mm(y) + (Mm(h) - ph) / 2)))
    ext.set("cx", str(pw))
    ext.set("cy", str(ph))
    slide.shapes._spTree.append(el_)
    return el_


def page_title(slide):
    for sh in slide.shapes:
        if sh.name == "Page title":
            return sh.text_frame.text.replace("\n", " ")
    return ""


# --------------------------------------------------------------------------- openers
def build_opener(prs, first, last, number, contact):
    """New dark page before a product line; returns the slide (appended at the end)."""
    src = prs.slides[first - 1]
    label = next(sh.text_frame.text for sh in src.shapes if sh.name == "Label")
    line_name = label.split("·")[-1].strip()
    dark = prs.slides[len(prs.slides) - 1]  # contact page: dark header/footer styles
    slide = prs.slides.add_slide(src.slide_layout)
    for ph in list(slide.placeholders):
        ph._element.getparent().remove(ph._element)
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor.from_string(NAVY)
    tree = slide.shapes._spTree
    for sh in dark.shapes:
        if sh.name in ("Running header", "Footer", "Page number"):
            tree.append(copy.deepcopy(sh._element))
    sec = copy.deepcopy(next(sh for sh in dark.shapes if sh.name == "Section label")._element)
    sec.find(".//" + qn("a:t")).text = line_name
    tree.append(sec)

    glow(slide, -20, 78, 250, 150)
    add_text(slide, 16, 24, 60, 30, f"{number:02d}", 64, SKY, bold=True)
    add_text(slide, 16, 55, 150, 5, f"Product line {number:02d}", 9, SKY, bold=True, caps=True, spc=120)
    add_text(slide, 16, 61, 178, 16, line_name, 34, "FFFFFF", bold=True)

    hero = max((sh for sh in src.shapes if sh.shape_type == 13 and not sh.name.startswith(("Glow", "Icon"))),
               key=lambda p: p.width * p.height)
    pic = copy_picture(hero, slide, 16, 92, 178, 118)
    set_effect(pic.find(qn("p:spPr")), shadow(22, 12, 55, "000000"))

    add_text(slide, 16, 219, 100, 5, "In this section", 9, SKY, bold=True, caps=True, spc=120)
    pages = list(range(first, last + 1))
    per_col = (len(pages) + 1) // 2
    for i, pno in enumerate(pages):
        x = 16 + (i // per_col) * 90
        y = 226 + (i % per_col) * 7.5
        add_text(slide, x, y, 12, 6, "", 11, SKY, bold=True).name = f"Section page {pno}"
        add_text(slide, x + 12, y, 76, 6, page_title(prs.slides[pno - 1]), 11, "FFFFFF")

    box = add_box(slide, 16, 260, 178, 17, "16314B", radius=0.18)
    box.name = "Call to action"
    add_text(slide, 22, 260, 100, 17, CTA_TITLE, 11.5, "FFFFFF", bold=True, anchor="ctr")
    add_text(slide, 110, 260, 78, 17, contact, 9.5, SKY, align="r", anchor="ctr")
    return slide


# --------------------------------------------------------------------------- restyle
def restyle_slide(slide, si, n_slides, dark):
    shapes = list(slide.shapes)
    header = next((s for s in shapes if s.name == "Header block"), None)
    if not dark:
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor.from_string(PAGE_BG)

    for sh in shapes:
        n = sh.name
        if "图片" in n:
            sh.name = "Picture"
        if getattr(sh, "has_table", False) and sh.has_table:
            for ri, row in enumerate(sh.table.rows):
                for cell in row.cells:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor.from_string("FFFFFF" if ri % 2 == 0 else "EAF0F6")
        spPr = sh._element.find(qn("p:spPr"))
        if spPr is not None:
            if n in STAGES:
                set_fill(spPr, gradient("FFFFFF", "D2DBE6"))
                set_effect(spPr, shadow(18, 7, 22))
            elif n in CARDS or n in WHITE_PANELS:
                set_fill(spPr, solid("FFFFFF"))
                set_line(spPr, "DCE3EB")
                set_effect(spPr, shadow(14, 5, 18))
            elif n.startswith("Chip") and not dark and spPr.find(qn("a:solidFill")) is not None \
                    and spPr.find(qn("a:solidFill"))[0].get("val") == "F3F6F9":
                set_fill(spPr, solid("FFFFFF"))
            elif n.startswith("Port marker"):
                set_fill(spPr, solid(ACCENT))
            elif sh.shape_type == 13 and not n.startswith(("Glow", "Icon")) and sh.width > Mm(25) \
                    and si != n_slides:
                if is_cutout(sh):
                    set_effect(spPr, shadow(10, 8, 40))
                else:
                    geom = spPr.find(qn("a:prstGeom"))
                    geom.set("prst", "roundRect")
                    av = geom.find(qn("a:avLst"))
                    av.clear() if av is not None else geom.append(el("a:avLst"))
                    (av if av is not None else geom.find(qn("a:avLst"))).append(el("a:gd", name="adj", fmla="val 4000"))
                    set_effect(spPr, shadow(14, 7, 30))

        if sh.has_text_frame:
            for r in sh._element.iter(qn("a:rPr")):
                sz = int(r.get("sz", "0"))
                if n in MIN_SIZE and 0 < sz < MIN_SIZE[n]:
                    r.set("sz", str(MIN_SIZE[n]))

    # Icons on stat cards.
    cards = [s for s in shapes if s.name in ("Card", "Fact tile")]
    for v in [s for s in shapes if s.name in ("Card value", "Fact value")]:
        t = v.text_frame.text
        icon = "city" if t.strip() == "3" else next((i for k, i in ICON_FOR if k in t), None)
        card = next((c for c in cards if c.left <= v.left < c.left + c.width
                     and c.top <= v.top < c.top + c.height), None)
        if icon and card is not None:
            d = Mm(6.5) if card.width > Mm(45) else Mm(5.5)
            slide.shapes.add_picture(os.path.join(ICONS, icon + ".png"),
                                     card.left + card.width - d - Mm(3), card.top + Mm(3), d, d).name = f"Icon {icon}"

    stage = next((s for s in shapes if s.name == "Product stage"), None)
    if header is not None and stage is not None:
        glow(slide, 5, 22, 200, 80, header._element)
    if si == 1 and stage is not None:
        glow(slide, -40, 50, 290, 250)
    if si == n_slides:
        glow(slide, 90, 140, 160, 150)

    for sp in slide.shapes._spTree.iter(qn("p:sp")):
        lst = sp.find(qn("p:spPr") + "/" + qn("a:effectLst"))
        ref = sp.find(".//" + qn("a:effectRef"))
        if ref is not None and (lst is None or len(lst) == 0):
            ref.set("idx", "0")


def fix_terms(prs):
    for s in prs.slides:
        for t in s.shapes._spTree.iter(qn("a:t")):
            if t.text:
                for old, new in TERMS:
                    t.text = t.text.replace(old, new)


def main(src, dst):
    prs = Presentation(src)
    n0 = len(prs.slides)
    contact_slide = prs.slides[n0 - 1]
    phone = next(sh.text_frame.text for sh in contact_slide.shapes if sh.name == "Phone").split("\n")[0]
    email = next(sh.text_frame.text for sh in contact_slide.shapes if sh.name.startswith("Email")).split("\n")[0] \
        if any(sh.name.startswith("Email") for sh in contact_slide.shapes) else "Admin@xyc-ltd.com"
    contact = f"{email.strip()}  ·  {phone.strip()}"

    # Restyle existing pages first (dark: cover and contact).
    for si, slide in enumerate(prs.slides, 1):
        restyle_slide(slide, si, n0, dark=si in (1, n0))

    # Build openers (appended), then move each before its section.
    openers = [build_opener(prs, first, last, i + 1, contact) for i, (first, last) in enumerate(SECTIONS)]
    lst = prs.slides._sldIdLst
    ids = list(lst)
    for opener, (first, _) in zip(openers, SECTIONS):
        rid = next(r.rId for r in prs.part.rels.values() if r._target is opener.part)
        e = next(x for x in lst if x.get(qn("r:id")) == rid)
        lst.remove(e)
        target = ids[first - 1]
        target.addprevious(e)

    # Final page numbers.
    order = list(prs.slides)
    final = {id(s): i for i, s in enumerate(order, 1)}
    old_index = {}  # original page number -> final page number
    k = 0
    for i, s in enumerate(order, 1):
        if s in openers:
            continue
        k += 1
        old_index[k] = i
    for s in order:
        for sh in s.shapes:
            if sh.name.startswith("Section page "):
                pno = old_index[int(sh.name.split()[-1])]
                sh.text_frame.paragraphs[0].runs[0].text = f"{pno:02d}"
        for f in s.shapes._spTree.iter(qn("a:fld")):
            if f.get("type") == "slidenum":
                f.find(qn("a:t")).text = str(final[id(s)])
    refs = {}
    for i, (first, last) in enumerate(SECTIONS):
        # page 4 refs currently read the pre-opener ranges
        refs[f"Page {first:02d}–{last:02d}"] = f"Page {old_index[first] - 1:02d}–{old_index[last]:02d}"
    mb = old_index[7]  # mainboard photo page
    for t in order[3].shapes._spTree.iter(qn("a:t")):
        if t.text in refs:
            t.text = refs[t.text]
    for sh in order[old_index[6] - 1].shapes:
        if sh.name == "Label" and sh.text_frame.text == "Mainboard options":
            sh.text_frame.paragraphs[0].runs[0].text = f"Mainboard options  ·  board photos on page {mb:02d}"
    fix_terms(prs)
    prs.save(dst)
    print("saved", dst, len(prs.slides), "pages")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
