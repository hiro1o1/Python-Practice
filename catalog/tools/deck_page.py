"""Insert the "Mainboard Options" page (after page 6) into the product catalog PPTX.

    python catalog/tools/deck_page.py "<in.pptx>" "<out.pptx>"

Every shape on the new page is cloned from an existing shape on page 6
("Configurations"), so fonts, colours and corner radii match the catalog.
Board pictures are the transparent cutouts written by build.py; port markers
are native shapes placed from the coordinates in boards.py.

Existing pages are only touched where the insert shifts page numbers: the
cached text of each slide-number field, and the fixed "Page NN–NN" references
on page 4.
"""

import copy
import json
import os
import sys

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.util import Emu, Mm, Pt

sys.path.insert(0, os.path.dirname(__file__))
from boards import BOARDS  # noqa: E402  (board names)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_SLIDE = 6       # 1-based; the new page copies its styles
INSERT_AFTER = 6

INK, BODY, MUTED, TEAL = "0B1B2B", "1B2733", "566270", "09678F"

# Port types in marker order. Simplified boards use 1–7; full boards add 8–9.
PORTS = ["HDMI in", "HDMI out", "LAN", "USB", "USB-C", "Audio in / out", "RS232", "DP in", "VGA in"]

# External ports marked on each board, in source-photo pixels (the same
# coordinates as the matching callouts in boards.py). Internal pin headers
# are left out.
PORT_POINTS = {
    "rk3576-a": {
        "HDMI in": [(617, 1098)], "HDMI out": [(745, 1097)], "LAN": [(1303, 1098)],
        "USB": [(948, 1097), (1065, 1098), (1182, 1098)], "USB-C": [(849, 1097)],
        "Audio in / out": [(1406, 1098), (1477, 1098)], "RS232": [(453, 555)],
    },
    "rk3588-a": {
        "HDMI in": [(1200, 686)], "HDMI out": [(1205, 583)], "LAN": [(1210, 477)],
        "USB": [(1215, 225), (1215, 288), (175, 129)], "USB-C": [(1190, 378)],
        "Audio in / out": [(548, 870), (620, 876)], "RS232": [(824, 896)],
    },
    "rk3576-b": {
        "HDMI in": [(426, 988), (500, 990)], "HDMI out": [(647, 991)], "LAN": [(873, 992), (955, 990)],
        "USB": [(1044, 212), (719, 991), (796, 989), (1146, 253)], "USB-C": [(396, 733)],
        "Audio in / out": [(1038, 988)], "RS232": [(405, 311)],
        "DP in": [(402, 663)], "VGA in": [(409, 550)],
    },
    "rk3588-b": {
        "HDMI in": [(1560, 707), (1562, 900), (1570, 1090)], "HDMI out": [(712, 1160)],
        "LAN": [(318, 1170), (410, 1168)],
        "USB": [(1560, 194), (1562, 248), (1558, 306), (280, 162), (276, 218), (280, 278)],
        "USB-C": [(1555, 400)], "Audio in / out": [(497, 1175), (577, 1182)], "RS232": [(1200, 1195)],
        "DP in": [(1552, 502)], "VGA in": [(1015, 1185)],
    },
}

COLUMNS = [
    {
        "title": "Simplified board",
        "note": "Best for rooms with one main source.",
        "boards": ["rk3576-a", "rk3588-a"],
    },
    {
        "title": "Full board",
        "note": "Best for rooms that connect several sources.",
        "boards": ["rk3576-b", "rk3588-b"],
    },
]

# Legend wording per board, where it differs from the plain port name.
LEGEND = {
    "rk3576-a": {"HDMI in": "1 × HDMI in", "HDMI out": "1 × HDMI out", "LAN": "1 × LAN"},
    "rk3588-a": {"HDMI in": "1 × HDMI in", "HDMI out": "1 × HDMI out", "LAN": "1 × LAN"},
    "rk3576-b": {"HDMI in": "2 × HDMI in", "HDMI out": "1 × HDMI out", "LAN": "2 × LAN",
                 "USB-C": "USB-C, 65 W"},
    "rk3588-b": {"HDMI in": "3 × HDMI in", "HDMI out": "1 × HDMI out", "LAN": "2 × LAN",
                 "USB-C": "USB-C in"},
}

PAGE_REFS = {"Page 05–10": "Page 05–11", "Page 11–13": "Page 12–14",
             "Page 14–15": "Page 15–16", "Page 16–18": "Page 17–19"}


# --------------------------------------------------------------------------- helpers
def templates(slide):
    found = {}
    for el in slide.shapes._spTree:
        nv = el.find(".//" + qn("p:cNvPr"))
        if nv is not None and nv.get("name") not in found:
            found[nv.get("name")] = el
    return found


class Page:
    def __init__(self, slide, tpl):
        self.slide, self.tpl = slide, tpl
        self.tree = slide.shapes._spTree
        self.next_id = 2

    def _id(self):
        self.next_id += 1
        return self.next_id

    def clone(self, name, x=None, y=None, w=None, h=None, text=None, rename=None):
        el = copy.deepcopy(self.tpl[name])
        nv = el.find(".//" + qn("p:cNvPr"))
        nv.set("id", str(self._id()))
        if rename:
            nv.set("name", rename)
        xfrm = el.find(".//" + qn("a:xfrm"))
        off, ext = xfrm.find(qn("a:off")), xfrm.find(qn("a:ext"))
        if x is not None:
            off.set("x", str(int(Mm(x)))); off.set("y", str(int(Mm(y))))
            ext.set("cx", str(int(Mm(w)))); ext.set("cy", str(int(Mm(h))))
        if text is not None:
            paras = el.findall(".//" + qn("a:p"))
            for p in paras[1:]:
                p.getparent().remove(p)
            runs = paras[0].findall(qn("a:r"))
            for r in runs[1:]:
                paras[0].remove(r)
            runs[0].find(qn("a:t")).text = text
        self.tree.append(el)
        return el

    def picture(self, path, x, y, w, h, name):
        pic = self.slide.shapes.add_picture(path, Mm(x), Mm(y), Mm(w), Mm(h))
        pic.name = name
        pic._element.find(".//" + qn("p:cNvPr")).set("descr", name)
        return pic

    def marker(self, cx, cy, n, d=3.0):
        s = self.slide.shapes.add_shape(MSO_SHAPE.OVAL, Mm(cx - d / 2), Mm(cy - d / 2), Mm(d), Mm(d))
        s.name = f"Port marker {n}"
        s.shadow.inherit = False
        s.fill.solid(); s.fill.fore_color.rgb = rgb(INK)
        s.line.color.rgb = rgb("FFFFFF"); s.line.width = Pt(0.75)
        tf = s.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.word_wrap = False
        tf.vertical_anchor = 3  # middle
        p = tf.paragraphs[0]
        p.alignment = 2  # center
        r = p.add_run(); r.text = str(n)
        r.font.size = Pt(6.5); r.font.name = "Segoe UI Semibold"; r.font.color.rgb = rgb("FFFFFF")
        return s


def rgb(h):
    from pptx.dml.color import RGBColor
    return RGBColor.from_string(h)


def fit(iw, ih, bw, bh):
    s = min(bw / iw, bh / ih)
    return iw * s, ih * s


# --------------------------------------------------------------------------- page
def build_page(prs):
    tslide = prs.slides[TEMPLATE_SLIDE - 1]
    tpl = templates(tslide)
    slide = prs.slides.add_slide(tslide.slide_layout)
    for ph in list(slide.placeholders):
        ph._element.getparent().remove(ph._element)
    pg = Page(slide, tpl)

    manifest = json.load(open(os.path.join(ROOT, "images", "cutouts.json")))
    boards = {b["id"]: b for b in BOARDS}

    for name in ("Running header", "Section label", "Footer", "Page number"):
        pg.clone(name)
    pg.clone("Page title", text="Mainboard Options")
    pg.clone("Subtitle", text="E5 series  ·  RK3576 and RK3588, each with a simplified or a full mainboard")

    col_x, col_w = (16, 108), 86
    stage_y, stage_h, stage_gap = (69, 175), 102, 4
    for ci, col in enumerate(COLUMNS):
        x = col_x[ci]
        pg.clone("Mainboard card", x, 48, col_w, 17)
        pg.clone("Mainboard title", x + 5, 51.5, col_w - 10, 6, text=col["title"])
        pg.clone("Mainboard note", x + 5, 57.5, col_w - 10, 5, text=col["note"])

        for bi, bid in enumerate(col["boards"]):
            b, m = boards[bid], manifest[bid]
            sy = stage_y[bi]
            pg.clone("Product stage", x, sy, col_w, stage_h)
            pg.clone("Caption", x + 5, sy + 4.5, 40, 4, text=b["name"].split()[0], rename="Platform label")
            para = pg.tree[-1].find(".//" + qn("a:pPr")); para.set("algn", "l")

            # Picture, fitted into the image area of the stage.
            ax, ay, aw, ah = x + 4, sy + 11, col_w - 8, 58
            pw, ph = fit(*m["size"], aw, ah)
            px, py = ax + (aw - pw) / 2, ay + (ah - ph) / 2
            pg.picture(os.path.join(ROOT, "images", f"{bid}-cutout.png"), px, py, pw, ph,
                       f"{b['name'].split()[0]} {col['title'].lower()}")
            k = pw / m["size"][0]  # mm per cutout pixel

            def to_mm(pt):
                return (px + (pt[0] - m["origin"][0]) * m["scale"] * k,
                        py + (pt[1] - m["origin"][1]) * m["scale"] * k)

            used = [p for p in PORTS if p in PORT_POINTS[bid]]
            for port in used:
                for pt in PORT_POINTS[bid][port]:
                    pg.marker(*to_mm(pt), PORTS.index(port) + 1)

            # Legend: two columns under the picture.
            ly, row_h = sy + 72, 5.4
            per_col = (len(used) + 1) // 2
            for i, port in enumerate(used):
                lx = x + 6 + (i // per_col) * 40
                yy = ly + (i % per_col) * row_h
                pg.marker(lx + 1.5, yy + 2.2, PORTS.index(port) + 1)
                pg.clone("Mainboard note", lx + 4.8, yy + 0.2, 34, 4.4,
                         text=LEGEND[bid].get(port, port), rename="Port legend")
                run = pg.tree[-1].find(".//" + qn("a:rPr"))
                run.find(qn("a:solidFill")).find(qn("a:srgbClr")).set("val", BODY)
    # The theme's effect style adds a drop shadow in some renderers even with an
    # empty effect list; page 6 renders flat in PowerPoint, so pin that here.
    for ref in slide.shapes._spTree.iter(qn("a:effectRef")):
        ref.set("idx", "0")
    return slide


def renumber(prs):
    """Refresh cached slide-number text and fixed page references."""
    for i, s in enumerate(prs.slides, 1):
        for el in s.shapes._spTree.iter(qn("a:fld")):
            if el.get("type") == "slidenum":
                el.find(qn("a:t")).text = str(i)
        for t in s.shapes._spTree.iter(qn("a:t")):
            if t.text in PAGE_REFS:
                t.text = PAGE_REFS[t.text]


def main(src, dst):
    prs = Presentation(src)
    slide = build_page(prs)
    rid = next(r.rId for r in prs.part.rels.values() if r._target is slide.part)
    lst = prs.slides._sldIdLst
    el = next(e for e in lst if e.get(qn("r:id")) == rid)
    lst.remove(el)
    lst.insert(INSERT_AFTER, el)
    renumber(prs)
    prs.save(dst)
    print("saved", dst)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
