"""Add product line 05 "Smart Retail & AI Devices" to the catalog PPTX.

    python catalog/tools/add_line05.py <in.pptx> <out.pptx>

Adds, after the Portable Smart TV pages:
  opener 05, then three pages for the 10.1" Self-service Price Checker and
  three for the Dual-Screen AI Translator.
Every shape is cloned from an existing page (header band, stages, cards,
chips, labels, tables), so fonts, colours and spacing match the catalog.
On page 4 the Outdoor card (03) becomes the line 05 card. Nothing else in
the existing pages changes apart from page numbers after the insert.

Sources: price checker spec sheet 10.1Machine.doc (XYC-CJ-D5) and the
supplier product images; translator data from the supplier slides (T10 Max).
"""

import copy
import os
import sys

from PIL import Image, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.util import Mm, Pt
from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
PHOTOS = os.path.join(os.path.dirname(HERE), "photos")
ICONS = os.path.join(HERE, "icons")
SEMIBOLD = os.path.expanduser("~/.fonts/SegoeUI-Semibold-subset.ttf")
REGULAR = os.path.expanduser("~/.fonts/SegoeUI-subset.ttf")

LINE = "Smart Retail & AI Devices"
PC = '10.1" Self-service Price Checker'
TR = "Dual-Screen AI Translator"


# --------------------------------------------------------------------------- helpers
def by_title(prs, title):
    for s in prs.slides:
        for sh in s.shapes:
            if sh.name in ("Page title", "TextBox 8") and sh.has_text_frame \
                    and sh.text_frame.text.replace("\n", " ") == title:
                return s
    raise KeyError(title)


def shape(slide, name, text=None, nth=0):
    hits = [sh for sh in slide.shapes if sh.name == name
            and (text is None or (sh.has_text_frame and sh.text_frame.text.startswith(text)))]
    return hits[nth]


def set_text(el, text):
    lines = text if isinstance(text, list) else [text]
    body = el if el.tag == qn("p:txBody") or el.tag == qn("a:txBody") else el.find(".//" + qn("p:txBody"))
    if body is None:
        body = el.find(".//" + qn("a:txBody"))
    paras = body.findall(qn("a:p"))
    tpl = paras[0]
    for p in paras:
        body.remove(p)
    for line in lines:
        p = copy.deepcopy(tpl)
        runs = p.findall(qn("a:r"))
        for r in runs[1:]:
            p.remove(r)
        runs[0].find(qn("a:t")).text = line
        body.append(p)


def width_mm(text, pt, bold=True):
    f = ImageFont.truetype(SEMIBOLD if bold else REGULAR, 1000)
    return f.getlength(text) / 1000 * pt * 25.4 / 72


def run_size(el):
    r = el.find(".//" + qn("a:rPr"))
    return int(r.get("sz", "1100")) / 100


class Page:
    def __init__(self, prs, layout_src, dark=False):
        self.prs = prs
        self.slide = prs.slides.add_slide(layout_src.slide_layout)
        for ph in list(self.slide.placeholders):
            ph._element.getparent().remove(ph._element)
        bg = layout_src.background.fill
        self.slide.background.fill.solid()
        self.slide.background.fill.fore_color.rgb = bg.fore_color.rgb
        self.tree = self.slide.shapes._spTree
        self.nid = 200

    def clone(self, src, x=None, y=None, w=None, h=None, text=None, name=None):
        el = copy.deepcopy(src._element)
        self.nid += 1
        nv = el.find(".//" + qn("p:cNvPr"))
        nv.set("id", str(self.nid))
        if name:
            nv.set("name", name)
        xfrm = el.find(".//" + qn("a:xfrm"))
        if xfrm is None:
            xfrm = el.find(".//" + qn("p:xfrm"))
        off, ext = xfrm.find(qn("a:off")), xfrm.find(qn("a:ext"))
        if x is not None:
            off.set("x", str(int(Mm(x))))
            off.set("y", str(int(Mm(y))))
        if w is not None:
            ext.set("cx", str(int(Mm(w))))
            ext.set("cy", str(int(Mm(h))))
        if text is not None:
            set_text(el, text)
        if src.shape_type == 13:  # picture: bring the image relationship across
            blip = el.find(".//" + qn("a:blip"))
            part = src.part.related_part(src._element.blipFill.blip.rEmbed)
            rid = self.slide.part.relate_to(part, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
            blip.set(qn("r:embed"), rid)
        self.tree.append(el)
        return el

    def picture(self, fname, x, y, w, h, alt, fit=True, shadow=None, rounded=False, anchor="c"):
        path = fname if os.path.isabs(fname) else os.path.join(PHOTOS, fname)
        iw, ih = Image.open(path).size
        if fit:
            s = min(w / iw, h / ih)
            pw, ph = iw * s, ih * s
            px = x + (w - pw) / 2
            py = y + (h - ph) / 2 if anchor == "c" else y + h - ph
            pic = self.slide.shapes.add_picture(path, Mm(px), Mm(py), Mm(pw), Mm(ph))
        else:  # fill the box, cropping the overflow
            pic = self.slide.shapes.add_picture(path, Mm(x), Mm(y), Mm(w), Mm(h))
            box, img = w / h, iw / ih
            if img > box:
                c = (1 - box / img) / 2
                pic.crop_left = pic.crop_right = c
            else:
                c = (1 - img / box) / 2
                pic.crop_top = pic.crop_bottom = c
        pic.name = alt
        pic._element.find(".//" + qn("p:cNvPr")).set("descr", alt)
        spPr = pic._element.spPr
        if rounded:
            geom = spPr.find(qn("a:prstGeom"))
            geom.set("prst", "roundRect")
            av = geom.find(qn("a:avLst"))
            av.append(etree.fromstring('<a:gd xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="adj" fmla="val 5000"/>'))
        if shadow:
            blur, dist, alpha = shadow
            spPr.append(etree.fromstring(
                '<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                f'<a:outerShdw blurRad="{int(blur*12700)}" dist="{int(dist*12700)}" dir="5400000" algn="t" rotWithShape="0">'
                f'<a:srgbClr val="0B1B2B"><a:alpha val="{int(alpha*1000)}"/></a:srgbClr></a:outerShdw></a:effectLst>'))
        return pic

    def icon(self, name, x, y, d=5.5):
        pic = self.slide.shapes.add_picture(os.path.join(ICONS, name + ".png"), Mm(x), Mm(y), Mm(d), Mm(d))
        pic.name = f"Icon {name}"
        return pic


class Kit:
    """Shapes to clone, collected from existing pages."""

    def __init__(self, prs):
        self.tv = by_title(prs, "Portable Smart Android TV")          # dark header page
        self.out = by_title(prs, "Ultra Durable Outdoor Weatherproof Advertising Displays")
        self.sig = by_title(prs, "4K Android Advertising Displays")
        self.spec = by_title(prs, "Platform & Specifications")        # light page
        self.ops = by_title(prs, "OPS PC Module")
        self.smart = by_title(prs, "4K AI Smartboards")
        self.opener = by_title(prs, "Portable Smart TV")
        s = self.out
        self.header_block = shape(s, "Header block")
        self.glow = shape(s, "Glow")
        self.stage = shape(s, "Product stage")
        self.dark_label = shape(s, "Label", "Product line")
        self.dark_title = shape(self.sig, "Page title")
        self.descriptor = shape(self.sig, "Descriptor")
        self.caption = shape(self.sig, "Caption")
        self.rule = shape(s, "Rule")
        self.spec_label = shape(s, "Spec label")
        self.spec_value = shape(s, "Spec value")
        self.label = shape(s, "Label", "Available sizes")
        self.chip = shape(s, "Chip 32\"")
        self.card = shape(s, "Card")
        self.card_value = shape(s, "Card value")
        self.card_note = shape(s, "Card note")
        self.features = shape(s, "Key features")
        self.size_note = shape(s, "Size note")
        self.light_title = shape(self.spec, "Page title")
        self.subtitle = shape(self.spec, "Subtitle")
        self.fig_tile = shape(self.spec, "Figure tile")
        self.figure = shape(self.spec, "Figure")
        self.fig_note = shape(self.spec, "Figure note")
        self.image_panel = shape(self.ops, "Image panel")
        self.table = next(sh for sh in self.ops.shapes if sh.has_table)
        self.use_title = shape(self.tv, "Use case")
        self.use_text = shape(self.tv, "Use case text")


def chrome(pg, kit, dark_src, section=None):
    for n in ("Running header", "Footer", "Page number"):
        pg.clone(shape(dark_src, n))
    if section is not None:
        pg.clone(shape(kit.spec, "Section label"), text=section)


def header_page(prs, kit, title, descriptor):
    pg = Page(prs, kit.out)
    pg.clone(kit.header_block)
    pg.clone(kit.glow)
    chrome(pg, kit, kit.out)
    pg.clone(kit.dark_label, text=f"Product line 05  ·  {LINE}")
    pg.clone(kit.dark_title, text=title)
    pg.clone(kit.descriptor, text=descriptor)
    return pg


def light_page(prs, kit, title, subtitle, section):
    pg = Page(prs, kit.spec)
    chrome(pg, kit, kit.spec, section)
    pg.clone(kit.light_title, text=title)
    pg.clone(kit.subtitle, text=subtitle)
    return pg


def chips(pg, kit, items, x, y, h=10.0, gap=2.6):
    size = run_size(kit.chip._element)
    for t in items:
        w = width_mm(t, size) + 9
        pg.clone(kit.chip, x, y, w, h, text=t, name=f"Chip {t}")
        x += w + gap
    return x


def spec_column(pg, kit, items, x, y, w, step=19):
    for label, value in items:
        pg.clone(kit.rule, x, y, w, 0.3)
        pg.clone(kit.spec_label, x, y + 3.2, w, 4, text=label)
        pg.clone(kit.spec_value, x, y + 9, w, 7.9, text=value)
        y += step


def cards(pg, kit, items, x0, y, w, h, cols, gap=6, icon=True):
    for i, (value, note, ic) in enumerate(items):
        x = x0 + (i % cols) * (w + gap)
        yy = y + (i // cols) * (h + gap)
        pg.clone(kit.card, x, yy, w, h)
        pg.clone(kit.card_value, x + 6, yy + 5.5, w - 12 - (7 if icon and ic else 0), 9, text=value)
        pg.clone(kit.card_note, x + 6, yy + 15, w - 12, h - 18, text=note)
        if icon and ic:
            pg.icon(ic, x + w - 3 - 5.5, yy + 6.5)


def features(pg, kit, cols, y, h):
    for i, col in enumerate(cols):
        pg.clone(kit.features, 16 + i * 94, y, 84, h, text=col)


def table(pg, kit, rows, x, y, w, row_h=7.2, name="Specifications"):
    el = pg.clone(kit.table, x, y, w, row_h * len(rows), name=name)
    tbl = el.find(".//" + qn("a:tbl"))
    trs = tbl.findall(qn("a:tr"))
    while len(trs) < len(rows):
        new = copy.deepcopy(trs[len(trs) % 2 - 2])
        trs[-1].addnext(new)
        trs = tbl.findall(qn("a:tr"))
    for extra in trs[len(rows):]:
        tbl.remove(extra)
    trs = tbl.findall(qn("a:tr"))
    for i, (tr, (k, v)) in enumerate(zip(trs, rows)):
        tr.set("h", str(int(Mm(row_h))))
        cells = tr.findall(qn("a:tc"))
        set_text(cells[0].find(qn("a:txBody")), k)
        set_text(cells[1].find(qn("a:txBody")), v)
        fill = "FFFFFF" if i % 2 == 0 else "EAF0F6"
        for c in cells:
            tcPr = c.find(qn("a:tcPr"))
            for old in tcPr.findall(qn("a:solidFill")) + tcPr.findall(qn("a:noFill")):
                tcPr.remove(old)
            sf = etree.SubElement(tcPr, qn("a:solidFill"))
            etree.SubElement(sf, qn("a:srgbClr")).set("val", fill)
    grid = tbl.find(qn("a:tblGrid")).findall(qn("a:gridCol"))
    first = min(int(Mm(34)), int(Mm(w)) // 3)
    grid[0].set("w", str(first))
    grid[1].set("w", str(int(Mm(w)) - first))


def callout(pg, kit, label, tip, text_xy, align="l"):
    """Leader line from a text label to a point on a product picture."""
    tx, ty = text_xy
    lw = width_mm(label.upper(), 7.5) * 1.18 + 2
    x = tx if align == "l" else tx - lw
    pg.clone(kit.caption, x, ty, lw, 4, text=label, name="Callout")
    para = pg.tree[-1].find(".//" + qn("a:pPr"))
    para.set("algn", align)
    lx = tx + lw + 1 if align == "l" else tx - lw - 1
    line = pg.slide.shapes.add_connector(1, Mm(lx), Mm(ty + 2), Mm(tip[0]), Mm(tip[1]))
    line.line.color.rgb = RGBColor.from_string("09678F")
    line.line.width = Pt(0.75)
    line.name = "Callout line"
    dot = pg.slide.shapes.add_shape(9, Mm(tip[0] - 1.1), Mm(tip[1] - 1.1), Mm(2.2), Mm(2.2))
    dot.fill.solid()
    dot.fill.fore_color.rgb = RGBColor.from_string("0E9BD6")
    dot.line.color.rgb = RGBColor.from_string("FFFFFF")
    dot.line.width = Pt(0.75)
    dot.shadow.inherit = False
    dot.name = "Callout dot"


# --------------------------------------------------------------------------- price checker
def pc_pages(prs, kit):
    # 1. Header page
    pg = header_page(prs, kit, PC, "Barcode price checker for retail  ·  Model XYC-CJ-D5")
    pg.clone(kit.stage, 16, 62, 178, 88)
    pg.picture("price-checker-front.png", 22, 66, 58, 70, "Price checker, front view", shadow=(10, 8, 40))
    pg.picture("price-checker-scanner.png", 82, 70, 40, 62, "Price checker, scanner window", shadow=(10, 8, 40))
    spec_column(pg, kit, [("Screen", '10.1" IPS · 1280 × 800'), ("Scanning", "1D and 2D barcodes"),
                          ("Platform", "Android 10 · quad-core"), ("Protection", "IP54 dust and splash")],
                130, 67, 58)
    pg.clone(kit.caption, 22, 143, 100, 4, text="Front view  ·  scanner window")
    pg.clone(kit.label, 16, 155, 178, 4, text="Built for")
    chips(pg, kit, ["Supermarkets", "Shopping malls", "Convenience stores", "Warehouses"], 16, 161.5)
    pg.clone(kit.rule, 16, 177, 178, 0.3)
    pg.clone(kit.label, 16, 180.2, 178, 4, text="Why customers use it")
    cards(pg, kit, [("No missing price tags", "Scan any product to see its name, price, promotion and original price.", "barcode"),
                    ("Same price at checkout", "Linked to the store ERP / POS, so shelf and till always match.", "zap"),
                    ("Easy for everyone", "Touch or scan; voice price announcement for visually impaired shoppers.", "touch"),
                    ("For every shopper", "Multi-language display; multi-currency on some models.", "globe")],
          16, 186, 86, 25, 2, gap=5)
    for i, (fig, note) in enumerate([("360°", "Reads barcodes at any rotation"),
                                     ("25–110 mm", "EAN-13 reading distance")]):
        x = 16 + i * 92
        pg.clone(kit.fig_tile, x, 250, 86, 27)
        pg.clone(kit.figure, x + 6, 251.5, 74, 16, text=fig)
        pg.clone(kit.fig_note, x + 6, 268.5, 74, 6, text=note)
    # 2. In the store
    pg = light_page(prs, kit, "In the Store", "From the shelf to the back office", LINE)
    pg.clone(kit.label, 16, 50, 178, 4, text="How it works")
    steps = [("1", "Scan", "Hold the barcode under the scanner window. 1D and 2D codes are read."),
             ("2", "See", "Name, price, promotion and original price appear in real time."),
             ("3", "Decide", "Member price, points and coupons after a QR member login.")]
    for i, (n, t, d) in enumerate(steps):
        x = 16 + i * 60.7
        pg.clone(kit.fig_tile, x, 56, 56.6, 40)
        pg.clone(kit.figure, x + 5, 58, 20, 16, text=n)
        pg.clone(kit.card_value, x + 17, 62.5, 36, 9, text=t)
        pg.clone(kit.card_note, x + 5, 76, 47, 18, text=d)
    pg.clone(kit.label, 16, 104, 178, 4, text="On the shop floor")
    # three tiles sized to each photo's own shape (equal height, no cropping)
    H = 42.6
    tiles = [("pc-uc-restock.jpg", 1.912, "Staff restocking shelves", "Next to the shelves",
              "Promotions and details where shoppers decide."),
             ("pc-uc-wall.jpg", 1.029, "Price checker mounted on a store wall", "Wall-mounted",
              "Fixed with the included wall bracket."),
             ("pc-uc-shopper.jpg", 1.047, "Shopper checking a product", "Self-service checks",
              "Less waiting, more shopping time.")]
    x = 16
    for f, ratio, alt, title, note in tiles:
        w = H * ratio
        pg.picture(f, x, 110, w, H, alt, fit=False, rounded=True, shadow=(12, 5, 20))
        pg.clone(kit.spec_value, x, 110 + H + 3, w, 7, text=title)
        pg.clone(kit.card_note, x, 110 + H + 10.5, w, 10, text=note)
        x += w + 4
    pg.clone(kit.label, 16, 190, 178, 4, text="For the store team")
    cards(pg, kit, [("Promotions", "Buy-two-get-one and spend-and-save rules, member prices, points and coupons.", "capital"),
                    ("Product details", "Ingredients, usage and origin: ideal for electronics, cosmetics and baby products.", "monitor"),
                    ("Search statistics", "Top 10 searched products and peak search times, to plan tags and promotions.", "area"),
                    ("Remote maintenance", "Reports scanner or network faults automatically for remote repair.", "cpu")],
          16, 196, 86, 32, 2)
    # 3. Specifications
    pg = light_page(prs, kit, "Specifications & Options", 'Model XYC-CJ-D5  ·  10.1" price checker', LINE)
    pg.clone(kit.label, 16, 50, 178, 4, text="Specifications")
    table(pg, kit, [("Display", '10.1" IPS LCD, 1280 × 800, 250–300 cd/m², 60 Hz'),
                    ("Touch", "10-point touch, 2 mm tempered glass"),
                    ("Platform", "A133 quad-core Cortex-A53, 1.5 GHz · Android 10"),
                    ("Memory", "2 GB RAM, 32 GB storage"),
                    ("Scanner", "640 × 480 CMOS, white LED, buzzer prompt"),
                    ("Barcodes", "QR Code, EAN-13/8, UPC-A/E, Code 128, Code 39 and more"),
                    ("Connectivity", "Wi-Fi 2.4 GHz, Bluetooth 5.0, RJ45 LAN, USB 2.0 + 3.0"),
                    ("Power", "DC 12 V adapter, max. 15 W, standby < 0.5 W"),
                    ("Environment", "Operating 0–40 °C, storage −20–60 °C")],
          16, 56, 178, row_h=8.2)
    pg.clone(kit.label, 16, 138, 60, 4, text="Wall mounting")
    pg.clone(kit.image_panel, 16, 144, 56, 74)
    pg.picture("pc-wall-plate.png", 20, 148, 48, 62, "Wall mounting plate, dimensions in mm")
    pg.clone(kit.caption, 16, 220, 56, 4, text="Plate 185 × 240 mm")
    pg.clone(kit.label, 78, 138, 116, 4, text="Options on request")
    cards(pg, kit, [("Camera", "720P by default; 1080P available", "camera"),
                    ("NFC / RFID", "Card reading on request", "zap"),
                    ("PoE power", "Power over Ethernet instead of the adapter", "zap"),
                    ("Configuration", "Other chip, memory and storage", "chip")],
          78, 144, 55, 34, 2, gap=6)
    spec_column(pg, kit, [("Size  ·  net 1.2 kg", "267 × 272 × 89 mm")], 16, 234, 56.3)
    spec_column(pg, kit, [("Carton  ·  gross 1.5 kg", "325 × 325 × 135 mm")], 76.8, 234, 56.3)
    spec_column(pg, kit, [("In the box", "Adapter, wall bracket")], 137.7, 234, 56.3)
    pg.clone(kit.size_note, 16, 258, 178, 6,
             text="Warranty card included. Options add cost; please confirm the configuration with our sales team before ordering.")


# --------------------------------------------------------------------------- translator
LANGS = ["English", "Chinese", "Cantonese", "Japanese", "Korean", "Spanish", "Portuguese", "French", "German",
         "Italian", "Dutch", "Swedish", "Greek", "Czech", "Polish", "Romanian", "Bulgarian", "Russian",
         "Ukrainian", "Kazakh", "Uzbek", "Mongolian", "Turkish", "Arabic", "Persian", "Urdu", "Hindi",
         "Bengali", "Tamil", "Thai", "Vietnamese", "Malay", "Indonesian", "Filipino", "Swahili", "Hausa",
         "Traditional Chinese"]


def tr_pages(prs, kit):
    # 1. Header page
    pg = header_page(prs, kit, TR, "Two screens, one conversation  ·  Model T10 Max")
    pg.clone(kit.stage, 16, 62, 178, 88)
    pg.picture("translator-white-45.png", 20, 67, 74, 68, "AI translator, white", shadow=(10, 8, 40))
    pg.picture("translator-gray-45.png", 94, 76, 56, 58, "AI translator, gray", shadow=(10, 8, 40))
    spec_column(pg, kit, [("Screens", '2 × 10" touch'), ("Cameras", "2 × 1080P HD"),
                          ("Microphones", "Dual-mic array"), ("Network", "Wi-Fi")],
                154, 67, 36)
    pg.clone(kit.caption, 22, 143, 130, 4, text="White  ·  gray")
    pg.clone(kit.label, 16, 155, 178, 4, text="Colours")
    chips(pg, kit, ["White", "Black", "Gray"], 16, 161.5)
    pg.clone(kit.rule, 16, 177, 178, 0.3)
    pg.clone(kit.label, 16, 180.2, 178, 4, text="Why it works")
    cards(pg, kit, [("Face to face", "One screen for each person, so both sides read and speak naturally.", "users"),
                    ("AI large model", "Real-time voice translation and speech output, in 35+ languages.", "chip"),
                    ("Hears clearly", "Dual-microphone array for reliable voice recognition, even in noisy places.", "zap"),
                    ("Keeps improving", "Online server updates keep raising translation quality.", "globe")],
          16, 186, 86, 25, 2, gap=5)
    for i, (fig, note) in enumerate([("35+", "Languages"), ("180+", "Countries and regions"),
                                     ("Free", "Lifetime translation")]):
        x = 16 + i * 60.7
        pg.clone(kit.fig_tile, x, 250, 56.6, 27)
        pg.clone(kit.figure, x + 6, 251.5, 46, 16, text=fig)
        pg.clone(kit.fig_note, x + 6, 268.5, 46, 6, text=note)
    # 2. In use
    pg = light_page(prs, kit, "Where It Works", "Wherever two languages meet", LINE)
    scenes = [("tr-uc-hotel-front.jpg", "Hotel front desk"), ("tr-uc-airport.jpg", "Airport counters"),
              ("tr-uc-hospital.jpg", "Hospitals and clinics"), ("tr-uc-shopping-mall.jpg", "Shopping malls"),
              ("tr-uc-office-meeting.jpg", "Business meetings"), ("tr-uc-restaurant.jpg", "Restaurants")]
    for i, (f, cap) in enumerate(scenes):
        x = 16 + (i % 3) * 61.33
        y = 50 + (i // 3) * 66
        pg.picture(f, x, y, 55.3, 56, cap, fit=False, rounded=True, shadow=(12, 5, 20))
        pg.clone(kit.caption, x, y + 58, 55.3, 4, text=cap)
    pg.clone(kit.label, 16, 186, 178, 4, text="Smart features")
    cards(pg, kit, [("Your brand on screen", "Custom wallpaper: upload your company logo in one click.", "monitor"),
                    ("Translation history", "Export records, share with the team in one tap, secure cloud backup.", "storage"),
                    ("Wakes by itself", "Motion sensor wakes the device when a visitor walks up.", "eye"),
                    ("Memo mode", "Several memo templates; key translated content summarised automatically.", "memory")],
          16, 192, 86, 30, 2)
    # 3. Design & specifications
    pg = light_page(prs, kit, "Design & Specifications", "Model T10 Max", LINE)
    pg.clone(kit.image_panel, 16, 50, 178, 78)
    pic = pg.picture("translator-white-45.png", 62, 55, 86, 68, "AI translator, labelled parts")
    # part callouts (positions measured on the white 45° render)
    px, py, pw, ph = (Mm(62), Mm(55), Mm(86), Mm(68))
    iw, ih = Image.open(os.path.join(PHOTOS, "translator-white-45.png")).size
    s = min(pw / iw, ph / ih)
    ox = pic.left
    oy = pic.top

    def at(fx, fy):
        return ((ox + fx * iw * s) / 36000, (oy + fy * ih * s) / 36000)

    callout(pg, kit, "Microphone array", at(0.49, 0.07), (22, 58))
    callout(pg, kit, "Camera", at(0.74, 0.12), (188, 58), align="r")
    callout(pg, kit, '10" IPS touch screen', at(0.25, 0.45), (22, 90))
    callout(pg, kit, "Microphone switch", at(0.44, 0.84), (22, 115))
    callout(pg, kit, "Second screen on the back", at(0.92, 0.35), (188, 90), align="r")
    callout(pg, kit, "Base", at(0.86, 0.9), (188, 115), align="r")
    pg.clone(kit.label, 16, 134, 178, 4, text="35+ languages")
    size = 8.5
    x, y = 16, 140
    for lang in LANGS:
        w = width_mm(lang, size, bold=False) + 5.2
        if x + w > 194:
            x, y = 16, y + 6.4
        el = pg.clone(kit.chip, x, y, w, 5.2, text=lang, name=f"Chip {lang}")
        for r in el.iter(qn("a:rPr")):
            r.set("sz", str(int(size * 100)))
        x += w + 1.4
    ty = y + 11
    pg.clone(kit.label, 16, ty, 178, 4, text="Specifications")
    table(pg, kit, [("CPU", "Dual-core CPU × 2"),
                    ("Screens", 'Dual 10" IPS HD touch screens'),
                    ("Cameras", "2 × 1920 × 1080 HD"),
                    ("Audio", "Dual-microphone array, 1 speaker"),
                    ("Network", "Wi-Fi 2.4 GHz"),
                    ("Power", "DC 12 V / 1.5 A, plug-in use"),
                    ("Size", "273.6 × 96.5 × 213.5 mm"),
                    ("Environment", "−10–50 °C, 30–80% RH"),
                    ("Certifications", "GB 4943-2022, RoHS, CE, FCC, SRRC"),
                    ("In the box", "Translator, power supply, manual, gift box")],
          16, ty + 6, 178, row_h=8.4)


# --------------------------------------------------------------------------- opener + portfolio
def opener(prs, kit, entries):
    src = kit.opener
    pg = Page(prs, src)
    for sh in src.shapes:
        pg.clone(sh)
    s = pg.slide
    for sh in list(s.shapes):
        t = sh.text_frame.text if sh.has_text_frame else ""
        if t == "04":
            sh.text_frame.paragraphs[0].runs[0].text = "05"
        elif t == "Product line 04":
            sh.text_frame.paragraphs[0].runs[0].text = "Product line 05"
        elif t == "Portable Smart TV" and sh.name != "Section label":
            sh.text_frame.paragraphs[0].runs[0].text = LINE
        elif sh.name == "Section label":
            sh.text_frame.paragraphs[0].runs[0].text = LINE
        if sh.shape_type == 13 and not sh.name.startswith("Glow"):
            left, top, w, h = sh.left, sh.top, sh.width, sh.height
            sh._element.getparent().remove(sh._element)
    # hero: both products
    tmp = pg.picture("price-checker-front.png", 26, 100, 70, 100, "Price checker", shadow=(22, 12, 55))
    pg.picture("translator-white-45.png", 96, 112, 92, 90, "AI translator", shadow=(22, 12, 55))
    # "In this section": reuse number/title boxes, one row per entry
    nums = [sh for sh in s.shapes if sh.name.startswith("Section page ")]
    titles = [sh for sh in s.shapes if sh.has_text_frame and sh.name.startswith("TextBox")
              and any(abs(sh.top - n.top) < Mm(1) and sh.left > n.left for n in nums)]
    num_tpl, title_tpl = nums[0], min(titles, key=lambda t: (t.top, t.left))
    for sh in nums + titles:
        sh._element.getparent().remove(sh._element)
    for i, (pno, title) in enumerate(entries):
        col, row = divmod(i, 3)
        x, y = 16 + col * 90, 226 + row * 7.5
        pg.clone(num_tpl, x, y, 12, 6, text=f"{pno:02d}", name=f"Section page {pno}")
        pg.clone(title_tpl, x + 12, y, 76, 6, text=title)
    return pg.slide


def portfolio(prs, first, last):
    s = prs.slides[3]
    card = [sh for sh in s.shapes if sh.name == "Number" and sh.text_frame.text == "03"][0]
    top = card.top
    group = [sh for sh in s.shapes if Mm(165) <= sh.top <= Mm(275) and sh.left < Mm(105)]
    for sh in group:
        if sh.name == "Number":
            sh.text_frame.paragraphs[0].runs[0].text = "05"
        elif sh.name == "Card title":
            sh.text_frame.paragraphs[0].runs[0].text = LINE
        elif sh.name == "Card facts":
            paras = sh.text_frame.paragraphs
            lines = ['10.1" self-service price checker', "1D / 2D barcode scanning · Android 10",
                     "Dual-screen AI translator · 35+ languages"]
            body = sh.text_frame._txBody
            tpl = paras[0]._p
            for p in list(paras):
                body.remove(p._p)
            for line in lines:
                p = copy.deepcopy(tpl)
                for r in p.findall(qn("a:r"))[1:]:
                    p.remove(r)
                p.find(qn("a:r")).find(qn("a:t")).text = line
                body.append(p)
        elif sh.name == "Page reference":
            sh.text_frame.paragraphs[0].runs[0].text = f"Page {first:02d}–{last:02d}"
        elif sh.shape_type == 13:
            stage = next(x for x in s.shapes if x.name == "Card stage" and abs(x.top - sh.top) < Mm(8)
                         and x.left <= sh.left < x.left + x.width)
            path = os.path.join(PHOTOS, "price-checker-front.png")
            iw, ih = Image.open(path).size
            bw, bh = stage.width - Mm(14), stage.height - Mm(8)
            k = min(bw / iw, bh / ih)
            w, h = int(iw * k), int(ih * k)
            pic = s.shapes.add_picture(path, stage.left + (stage.width - w) // 2, stage.top + (stage.height - h) // 2, w, h)
            pic.name = "10.1 inch self-service price checker"
            pic._element.find(".//" + qn("p:cNvPr")).set("descr", pic.name)
            pic._element.spPr.append(copy.deepcopy(sh._element.spPr.find(qn("a:effectLst")))) \
                if sh._element.spPr.find(qn("a:effectLst")) is not None else None
            sh._element.addprevious(pic._element)
            sh._element.getparent().remove(sh._element)


def swap_production_photo(prs):
    """Production page: replace the black floor-standing photo with the factory photo, filling the same box."""
    s = by_title(prs, "Production by Product Line")
    old = next(sh for sh in s.shapes if sh.shape_type == 13 and sh.name.startswith("Black floor-standing"))
    path = os.path.join(PHOTOS, "factory-signage-floorstanding.jpg")
    iw, ih = Image.open(path).size
    pic = s.shapes.add_picture(path, old.left, old.top, old.width, old.height)
    box, img = old.width / old.height, iw / ih
    if img > box:
        c = (1 - box / img) / 2
        pic.crop_left = pic.crop_right = c
    else:  # taller photo: trim mostly ceiling, keep the floor and stands
        c = 1 - img / box
        pic.crop_top, pic.crop_bottom = c * 0.8, c * 0.2
    pic.name = "Floor-standing advertising displays in production"
    pic._element.find(".//" + qn("p:cNvPr")).set("descr", pic.name)
    for tag in ("a:prstGeom", "a:effectLst"):  # same rounded corners and shadow as before
        src = old._element.spPr.find(qn(tag))
        dst = pic._element.spPr.find(qn(tag))
        if src is not None:
            if dst is not None:
                pic._element.spPr.replace(dst, copy.deepcopy(src))
            else:
                pic._element.spPr.append(copy.deepcopy(src))
    old._element.addprevious(pic._element)
    old._element.getparent().remove(old._element)


# --------------------------------------------------------------------------- main
def main(src, dst):
    prs = Presentation(src)
    kit = Kit(prs)
    after = by_title(prs, "Platform & Specifications")
    n_before = len(prs.slides)
    pc_pages(prs, kit)
    tr_pages(prs, kit)
    new = list(prs.slides)[n_before:]
    pos = list(prs.slides).index(after) + 2  # page number of the opener
    titles = [PC, "In the Store", "Specifications & Options", TR, "Where It Works", "Design & Specifications"]
    op = opener(prs, kit, [(pos + 1 + i, t) for i, t in enumerate(titles)])
    # order: opener, then the six product pages, right after "Platform & Specifications"
    lst = prs.slides._sldIdLst
    rid_of = {r._target: r.rId for r in prs.part.rels.values()}
    anchor = next(e for e in lst if e.get(qn("r:id")) == rid_of[after.part])
    for sl in [op] + new:
        e = next(x for x in lst if x.get(qn("r:id")) == rid_of[sl.part])
        lst.remove(e)
        anchor.addnext(e)
        anchor = e
    for i, s in enumerate(prs.slides, 1):
        for f in s.shapes._spTree.iter(qn("a:fld")):
            if f.get("type") == "slidenum":
                f.find(qn("a:t")).text = str(i)
    portfolio(prs, pos, pos + 6)
    swap_production_photo(prs)
    prs.save(dst)
    print("saved", dst, len(prs.slides), "pages; line 05 on pages", pos, "to", pos + 6)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
