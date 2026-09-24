"""Catalog-wide visual refresh: depth, lighting and readability. No content is removed.

    python catalog/tools/restyle.py <in.pptx> <out.pptx>

* product stages get a top-lit gradient and a soft shadow
* grey cards and tiles become white, elevated cards
* product pictures get a soft contact shadow
* dark section openers and the cover get a glow behind the hero product
* stat cards get a line icon; spec tables get banded rows
* the smallest labels are enlarged and page references emphasised
"""

import os
import sys

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.util import Mm

ICONS = os.path.join(os.path.dirname(__file__), "icons")
A = "http://schemas.openxmlformats.org/drawingml/2006/main"

STAGES = {"Product stage", "Card stage", "Figure tile"}
CARDS = {"Card", "Use case", "Mainboard card", "Fact tile", "Statement panel"}
WHITE_PANELS = {"Image panel", "Series stage", "Drawing frame"}
SMALL_TEXT = {"Caption", "Label", "Card note", "Use case note", "Fact label", "Highlight label",
              "Series finish", "Platform label", "Spec label"}
ICON_FOR = [("Since", "calendar"), ("50–100", "users"), ("RMB", "capital"), ("m²", "area"),
            ("80+", "globe"), ("4K", "monitor"), ("points", "touch"), ("ms", "zap"), ("°", "eye"),
            ("Core", "cpu"), ("TB", "storage"), ("GB", "memory"), ("Air-cooled", "fan"),
            ("Air conditioner", "snow"), ("MTK", "chip"), ("MP", "camera")]
FACT_3 = "city"  # the "3" manufacturing-cities fact


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
    anchor = None
    for tag in ("a:ln", "a:solidFill", "a:gradFill", "a:noFill", "a:blipFill", "a:prstGeom", "a:custGeom"):
        anchor = spPr.find(qn(tag))
        if anchor is not None:
            break
    if anchor is None:
        spPr.append(lst)
    else:
        anchor.addnext(lst)


def set_fill(spPr, fill):
    for tag in ("a:solidFill", "a:gradFill", "a:noFill"):
        old = spPr.find(qn(tag))
        if old is not None:
            spPr.replace(old, fill)
            return


def gradient(top, bottom, path=None, stops=None):
    g = el("a:gradFill", rotWithShape=1)
    gs = el("a:gsLst")
    for pos, c in stops or [(0, color(top)), (100000, color(bottom))]:
        s = el("a:gs", pos=pos); s.append(c); gs.append(s)
    g.append(gs)
    if path:
        p = el("a:path", path=path)
        p.append(el("a:fillToRect", l=50000, t=50000, r=50000, b=50000))
        g.append(p)
    else:
        g.append(el("a:lin", ang=5400000, scaled=0))
    return g


def glow(slide, x, y, w, h, after=None):
    """Soft light behind the hero product (a transparent PNG renders the same everywhere)."""
    pic = slide.shapes.add_picture(os.path.join(ICONS, "glow.png"), Mm(x), Mm(y), Mm(w), Mm(h))
    pic.name = "Glow"
    tree = slide.shapes._spTree
    if after is not None:
        after.addnext(pic._element)
    else:
        tree.insert(2, pic._element)  # first shape: behind everything


def name(shape):
    return shape.name.split(" ")[0] if shape.name.startswith("Chip") else shape.name


def restyle(prs):
    for si, slide in enumerate(prs.slides, 1):
        shapes = list(slide.shapes)
        header = next((s for s in shapes if s.name == "Header block"), None)
        for sh in shapes:
            n = sh.name
            if "图片" in n:
                sh.name = "Picture"
            if getattr(sh, "has_table", False) and sh.has_table:
                for ri, row in enumerate(sh.table.rows):
                    if ri % 2 == 1:
                        for cell in row.cells:
                            cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor.from_string("F1F4F8")
            spPr = sh._element.find(qn("p:spPr"))
            if spPr is None:
                continue
            if n in STAGES:
                set_fill(spPr, gradient("F7F9FC", "E0E6EE"))
                set_effect(spPr, shadow(14, 4, 12))
            elif n in CARDS:
                set_fill(spPr, el("a:solidFill")); spPr.find(qn("a:solidFill")).append(color("FFFFFF"))
                ln = spPr.find(qn("a:ln"))
                new_ln = el("a:ln", w=6350); f = el("a:solidFill"); f.append(color("DFE5EC")); new_ln.append(f)
                spPr.replace(ln, new_ln) if ln is not None else spPr.append(new_ln)
                set_effect(spPr, shadow(10, 3, 10))
            elif n in WHITE_PANELS:
                set_effect(spPr, shadow(10, 3, 10))
            elif sh.shape_type == 13 and sh.width > Mm(25) and si not in (1, len(prs.slides)) \
                    and not n.startswith(("Glow", "Icon")):
                set_effect(spPr, shadow(9, 5, 22))
            elif sh.shape_type == 13 and si == 1:
                set_effect(spPr, shadow(14, 8, 35))

            if sh.has_text_frame:
                for r in sh._element.iter(qn("a:rPr")):
                    sz = int(r.get("sz", "0"))
                    if n in SMALL_TEXT and 0 < sz <= 800:
                        r.set("sz", str(sz + 50))
                    if n == "Page reference":
                        r.set("sz", "950")


        # Icons on stat cards: top-right corner of the card that holds the value.
        cards = [s for s in shapes if s.name in ("Card", "Fact tile")]
        for v in [s for s in shapes if s.name in ("Card value", "Fact value")]:
            t = v.text_frame.text
            icon = FACT_3 if t.strip() == "3" else next((i for k, i in ICON_FOR if k in t), None)
            card = next((c for c in cards if c.left <= v.left < c.left + c.width and c.top <= v.top < c.top + c.height), None)
            if not icon or card is None:
                continue
            d = Mm(5.5) if card.width > Mm(45) else Mm(4.5)
            top = card.top + Mm(3.2)
            pic = slide.shapes.add_picture(os.path.join(ICONS, icon + ".png"),
                                           card.left + card.width - d - Mm(3.2), top, d, d)
            pic.name = f"Icon {icon}"

        # Glow behind the hero product on dark pages.
        stage = next((s for s in shapes if s.name == "Product stage"), None)
        if header is not None and stage is not None:
            glow(slide, 15, 30, 180, 64, header._element)
        if si == 1 and stage is not None:
            glow(slide, -30, 60, 270, 230)

        # Theme effect refs add shadows in some renderers where none are set.
        for sp in slide.shapes._spTree.iter(qn("p:sp")):
            lst = sp.find(qn("p:spPr") + "/" + qn("a:effectLst"))
            ref = sp.find(".//" + qn("a:effectRef"))
            if ref is not None and (lst is None or len(lst) == 0):
                ref.set("idx", "0")


if __name__ == "__main__":
    prs = Presentation(sys.argv[1])
    restyle(prs)
    prs.save(sys.argv[2])
    print("saved", sys.argv[2])
