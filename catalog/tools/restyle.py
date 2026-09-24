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
# Opener hero: (page, picture-name prefix) so the opener does not repeat the
# picture on the page right after it. None = the section's own hero.
OPENER_HERO = {}
# The page after an opener shows a different picture of the same product, so
# the opener's hero is not repeated: base page -> (picture to replace, (page, source picture)).
# New photos for pages after an opener: base page -> (picture to replace, file in catalog/photos).
# The photo fills the page's product stage.
NEXT_PAGE_PHOTO = {5: ("Line-up of 4K AI smartboards", "smartboards-classroom.png",
                       "4K AI smartboards on mobile stands in a classroom")}
PHOTOS = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "photos")
# Pages after an opener that get a new product cut-out in the same box: base page -> (picture, file).
NEXT_PAGE_CUTOUT = {12: ("Wall-mounted advertising display, front", "wall-mounted-display.png",
                         "Wall-mounted advertising display, front view")}
NEXT_PAGE_SWAP = {17: ("Portable smart Android TV on a mobile", (18, "HS Series product render"))}
FONT_FILES = [os.path.expanduser("~/.fonts/SegoeUI-Semibold-subset.ttf"),
              os.path.expanduser("~/.fonts/SegoeUI-subset.ttf")]


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


def text_extent(shape):
    """(width of the widest line, height of the first line) in EMU."""
    from PIL import ImageFont
    widest, first_h = 0, None
    for p in shape.text_frame.paragraphs:
        text = "".join(r.text for r in p.runs)
        if not p.runs:
            continue
        rpr = p.runs[0]._r.find(qn("a:rPr"))
        size = int(rpr.get("sz", "1000")) / 100
        bold = "Semibold" in (p.runs[0].font.name or "")
        caps = rpr.get("cap") == "all"
        spc = int(rpr.get("spc", "0")) / 100
        if caps:
            text = text.upper()
        path = FONT_FILES[0 if bold else 1]
        try:
            f = ImageFont.truetype(path, 1000)
            w_pt = f.getlength(text) / 1000 * size
        except OSError:
            w_pt = len(text) * size * 0.56
        w_pt += spc * len(text)
        widest = max(widest, min(Pt(w_pt), shape.width))
        if first_h is None:
            first_h = Pt(size * 1.33)
    return widest, first_h or Pt(10)


def swap_picture(target, source):
    """Show ``source``'s image in ``target``'s place, fitted into the same box."""
    rid = target.part.relate_to(source.part.related_part(source._element.blipFill.blip.rEmbed),
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
    target._element.blipFill.blip.set(qn("r:embed"), rid)
    src = target._element.blipFill.find(qn("a:srcRect"))
    if src is not None:
        target._element.blipFill.remove(src)
    x, y, w, h = target.left, target.top, target.width, target.height
    s = min(w / source.width, h / source.height)
    nw, nh = int(source.width * s), int(source.height * s)
    target.left, target.top, target.width, target.height = x + (w - nw) // 2, y + (h - nh) // 2, nw, nh
    target.name = source.name


def place_photo(slide, old, path, descr):
    """Replace ``old`` with a photo that fills the product stage behind it (cropped, rounded)."""
    stage = next(sh for sh in slide.shapes if sh.name == "Product stage")
    x, y, w, h = stage.left, stage.top, stage.width, stage.height
    pic = slide.shapes.add_picture(path, x, y, w, h)
    iw, ih = Image.open(path).size
    box, img = w / h, iw / ih
    if img > box:
        c = (1 - box / img) / 2
        pic.crop_left = pic.crop_right = c
    else:
        c = (1 - img / box) / 2
        pic.crop_top = c * 0.6
        pic.crop_bottom = c * 1.4
    pic.name = descr
    pic._element.find(".//" + qn("p:cNvPr")).set("descr", descr)
    spPr = pic._element.spPr
    geom = spPr.find(qn("a:prstGeom"))
    geom.set("prst", "roundRect")
    av = geom.find(qn("a:avLst"))
    av.append(el("a:gd", name="adj", fmla="val 3846"))
    set_effect(spPr, shadow(18, 7, 22))
    old._element.addprevious(pic._element)
    old._element.getparent().remove(old._element)


PRODUCTION_AFTER = 20  # base page "Factory & Production"
PRODUCTION = [
    ("Interactive smartboards", [("factory-smartboards.jpg", "Smartboards on test racks in the factory")]),
    ("Advertising displays", [("factory-signage-white.jpg", "White floor-standing advertising displays in production"),
                              ("factory-signage-black.jpg", "Black floor-standing advertising displays in production")]),
    ("Portable smart TV", [("factory-smarttv-line.jpg", "Portable smart TV production line"),
                           ("factory-smarttv-office.jpg", "Portable smart TVs ready for dispatch")]),
]


def build_production_page(prs):
    src = prs.slides[PRODUCTION_AFTER - 1]
    slide = prs.slides.add_slide(src.slide_layout)
    for ph in list(slide.placeholders):
        ph._element.getparent().remove(ph._element)
    tree = slide.shapes._spTree
    for sh in src.shapes:
        if sh.name in ("Running header", "Section label", "Footer", "Page number", "Page title", "Lead"):
            e = copy.deepcopy(sh._element)
            tree.append(e)
    shapes = {sh.name: sh for sh in slide.shapes}
    shapes["Page title"].text_frame.paragraphs[0].runs[0].text = "Production by Product Line"
    lead = shapes["Lead"]
    lead.text_frame.paragraphs[0].runs[0].text = "Photos from our factory floor."
    for r in lead.text_frame.paragraphs[0].runs[1:]:
        r._r.getparent().remove(r._r)
    for p in lead.text_frame.paragraphs[1:]:
        p._p.getparent().remove(p._p)
    label_tpl = next(sh for sh in prs.slides[7].shapes if sh.name == "Label")  # "For Meeting Room" (light page)
    y = 48
    for title, photos in PRODUCTION:
        lab = copy.deepcopy(label_tpl._element)
        lab.find(".//" + qn("a:t")).text = title
        off = lab.find(".//" + qn("a:off")); ext = lab.find(".//" + qn("a:ext"))
        off.set("x", str(int(Mm(16)))); off.set("y", str(int(Mm(y))))
        ext.set("cx", str(int(Mm(178)))); ext.set("cy", str(int(Mm(4))))
        tree.append(lab)
        h = 68 if len(photos) == 1 else 58
        w = 178 if len(photos) == 1 else 86
        for i, (fname, descr) in enumerate(photos):
            pic = slide.shapes.add_picture(os.path.join(PHOTOS, fname), Mm(16 + i * 92), Mm(y + 6), Mm(w), Mm(h))
            pic.name = descr
            pic._element.find(".//" + qn("p:cNvPr")).set("descr", descr)
        y += 6 + h + 8
    restyle_slide(slide, 0, -1, dark=False)
    return slide


# --------------------------------------------------------------------------- image clean-up
def _defringe(rgba):
    """Pull edge colours from the opaque interior and trim the 1 px white fringe."""
    import cv2
    import numpy as np
    a = rgba[..., 3].astype(np.float32)
    rgb = rgba[..., :3].astype(np.float32)
    solid = (a >= 250).astype(np.float32)
    edge = (a > 0) & (a < 250)
    if edge.any() and solid.any():
        fill = rgb.copy()
        for k in (3, 7, 15):
            num = cv2.GaussianBlur(rgb * solid[..., None], (0, 0), k)
            den = cv2.GaussianBlur(solid, (0, 0), k)[..., None]
            est = num / np.maximum(den, 1e-4)
            need = edge & (den[..., 0] > 0.02)
            fill[need] = est[need]
            edge = edge & ~need
        rgb = np.where(((a > 0) & (a < 250))[..., None], fill, rgb)
    a = cv2.erode(a, np.ones((3, 3), np.uint8))
    a = cv2.GaussianBlur(a, (0, 0), 0.6)
    return np.dstack([np.clip(rgb, 0, 255), a]).astype(np.uint8)


def _knock_out_white(rgb):
    """White/near-white background connected to the border -> transparent (soft for light shadows)."""
    import cv2
    import numpy as np
    x = rgb.astype(np.int32)
    w = x.min(-1)
    sat = x.max(-1) - x.min(-1)
    cand = ((w >= 247) & (sat < 10)).astype(np.uint8)
    n, lab = cv2.connectedComponents(cand)
    border = set(np.unique(np.r_[lab[0], lab[-1], lab[:, 0], lab[:, -1]])) - {0}
    bg = np.isin(lab, list(border))
    alpha = np.full(w.shape, 255.0)
    alpha[bg] = np.clip((253 - w[bg]) / 6.0, 0, 1) * 255
    return _defringe(np.dstack([rgb, alpha.astype(np.uint8)]))


def clean_images(prs):
    """Make white-background product shots real cut-outs and defringe existing cut-outs."""
    import numpy as np
    done = {}
    for slide in prs.slides:
        for sh in list(slide.shapes):
            if sh.shape_type != 13 or sh.name.startswith(("Glow", "Icon")):
                continue
            blip = sh._element.blipFill.blip
            part = slide.part.related_part(blip.rEmbed)
            key = part.partname
            if key not in done:
                im = Image.open(io.BytesIO(part.blob))
                arr = np.asarray(im.convert("RGBA"))
                corners = [arr[0, 0], arr[0, -1], arr[-1, 0], arr[-1, -1]]
                out = None
                if im.mode == "RGBA" and arr[..., 3].min() < 200:
                    out = _defringe(arr)
                elif all(c[3] == 255 and c[:3].min() > 225 for c in corners):
                    out = _knock_out_white(arr[..., :3])
                if out is not None:
                    buf = io.BytesIO()
                    Image.fromarray(out, "RGBA").save(buf, "PNG", optimize=True)
                    done[key] = buf.getvalue()
                else:
                    done[key] = None
            if done[key] is not None:
                old = blip.rEmbed
                _, rid = slide.part.get_or_add_image_part(io.BytesIO(done[key]))
                blip.set(qn("r:embed"), rid)
                if old != rid and f'"{old}"' not in etree.tostring(slide._element).decode():
                    slide.part.drop_rel(old)


def fit_in_panels(slide):
    """Keep product pictures inside the white card they sit on (3 mm clearance)."""
    panels = [s for s in slide.shapes if s.name in ("Image panel", "Series stage")]
    for sh in slide.shapes:
        if sh.shape_type != 13 or sh.name.startswith(("Glow", "Icon")):
            continue
        cx = sh.left + sh.width // 2
        panel = next((p for p in panels if p.left <= cx <= p.left + p.width
                      and p.top <= sh.top < p.top + p.height), None)
        if panel is None:
            continue
        limit = panel.top + panel.height - Mm(3)
        if sh.top + sh.height > limit:
            k = (limit - sh.top) / sh.height
            nw, nh = int(sh.width * k), int(sh.height * k)
            sh.left, sh.width, sh.height = cx - nw // 2, nw, nh


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

    spec = OPENER_HERO.get(number)
    if spec:
        hero = next(sh for sh in prs.slides[spec[0] - 1].shapes
                    if sh.shape_type == 13 and sh.name.startswith(spec[1]))
    else:
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

    # Icons on stat cards: one row type per page (value line if every card has
    # room there, else the note line), right-aligned and centred on that line.
    cards = [s for s in shapes if s.name in ("Card", "Fact tile")]
    texts = [s for s in shapes if s.has_text_frame and s.name in ("Card value", "Fact value", "Card note", "Fact label")]
    placements = []
    for v in [s for s in shapes if s.name in ("Card value", "Fact value")]:
        t = v.text_frame.text
        icon = "city" if t.strip() == "3" else next((i for k, i in ICON_FOR if k in t), None)
        card = next((c for c in cards if c.left <= v.left < c.left + c.width
                     and c.top <= v.top < c.top + c.height), None)
        if not icon or card is None:
            continue
        inside = [x for x in texts if card.left <= x.left < card.left + card.width
                  and card.top <= x.top < card.top + card.height]
        note = next((x for x in inside if x.name in ("Card note", "Fact label")), None)
        right = card.left + card.width - Mm(3.5)
        rows = {}
        for key, row in (("value", v), ("note", note)):
            if row is not None:
                w, h = text_extent(row)
                rows[key] = (right - (row.left + w), row, h)
        placements.append((icon, right, rows, card.top))
    d = Mm(5.5)
    need = d + Mm(2.5)
    group_key = {}
    for top in {p[3] for p in placements}:
        group = [r for _, _, r, t in placements if t == top]
        options = [k for k in ("value", "note") if all(k in r for r in group)]
        group_key[top] = next((k for k in options if all(r[k][0] >= need for r in group)),
                              max(options, key=lambda k: min(r[k][0] for r in group)) if options else None)
    for icon, right, rows, top in placements:
        k = group_key[top] or max(rows, key=lambda kk: rows[kk][0])
        free, row, h = rows[k]
        size = d if free >= need else max(Mm(4), free - Mm(2))
        cy = row.top + h / 2
        pic = slide.shapes.add_picture(os.path.join(ICONS, icon + ".png"), right - size, int(cy - size / 2), size, size)
        pic.name = f"Icon {icon}"

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

    clean_images(prs)
    for slide in prs.slides:
        fit_in_panels(slide)
    # Restyle existing pages first (dark: cover and contact).
    for si, slide in enumerate(prs.slides, 1):
        restyle_slide(slide, si, n0, dark=si in (1, n0))

    # Build openers (appended), then move each before its section.
    openers = [build_opener(prs, first, last, i + 1, contact) for i, (first, last) in enumerate(SECTIONS)]
    production = build_production_page(prs)
    for page, (name, fname, descr) in NEXT_PAGE_PHOTO.items():
        slide = prs.slides[page - 1]
        old = next(sh for sh in slide.shapes if sh.shape_type == 13 and sh.name.startswith(name))
        place_photo(slide, old, os.path.join(PHOTOS, fname), descr)
    for page, (name, fname, descr) in NEXT_PAGE_CUTOUT.items():
        slide = prs.slides[page - 1]
        old = next(sh for sh in slide.shapes if sh.shape_type == 13 and sh.name.startswith(name))
        path = os.path.join(PHOTOS, fname)
        iw, ih = Image.open(path).size
        sc = min(old.width / iw, old.height / ih)
        w, h = int(iw * sc), int(ih * sc)
        pic = slide.shapes.add_picture(path, old.left + (old.width - w) // 2, old.top + (old.height - h) // 2, w, h)
        pic.name = descr
        pic._element.find(".//" + qn("p:cNvPr")).set("descr", descr)
        set_effect(pic._element.spPr, shadow(10, 8, 40))
        old._element.addprevious(pic._element)
        old._element.getparent().remove(old._element)
    for page, (name, (src_page, src_name)) in NEXT_PAGE_SWAP.items():
        target = next(sh for sh in prs.slides[page - 1].shapes if sh.shape_type == 13 and sh.name.startswith(name))
        source = next(sh for sh in prs.slides[src_page - 1].shapes if sh.shape_type == 13 and sh.name.startswith(src_name))
        swap_picture(target, source)
    lst = prs.slides._sldIdLst
    ids = list(lst)
    for opener, (first, _) in zip(openers, SECTIONS):
        rid = next(r.rId for r in prs.part.rels.values() if r._target is opener.part)
        e = next(x for x in lst if x.get(qn("r:id")) == rid)
        lst.remove(e)
        target = ids[first - 1]
        target.addprevious(e)
    rid = next(r.rId for r in prs.part.rels.values() if r._target is production.part)
    e = next(x for x in lst if x.get(qn("r:id")) == rid)
    lst.remove(e)
    ids[PRODUCTION_AFTER - 1].addnext(e)

    # Final page numbers.
    order = list(prs.slides)
    final = {id(s): i for i, s in enumerate(order, 1)}
    old_index = {}  # original page number -> final page number
    k = 0
    for i, s in enumerate(order, 1):
        if s in openers or s is production:
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
