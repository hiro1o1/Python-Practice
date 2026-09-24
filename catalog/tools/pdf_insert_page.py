"""Insert one new page into the PowerPoint-exported catalog PDF, keeping every
existing page exactly as exported.

    python catalog/tools/pdf_insert_page.py <catalog.pdf> <new-deck.pdf> <page> <fonts-dir> <out.pdf>

<new-deck.pdf> is an export of the updated PPTX; its page <page> is inserted
at that position. On the existing pages only numbers are rewritten, in the
same embedded Segoe UI fonts, size and colour:
* printed page numbers after the insert
* "Page NN–NN" references on the portfolio page
* "In this section" numbers on the section openers, plus one new line for
  the inserted page on its own opener
"""

import glob
import sys

import pymupdf

NEW_TITLE = '10.1" Self-service Price Checker'
FOOTER_Y = 800


def spans(page):
    for b in page.get_text("rawdict")["blocks"]:
        for line in b.get("lines", []):
            for s in line["spans"]:
                yield s, "".join(c["c"] for c in s["chars"])


def rgb(c):
    return ((c >> 16) & 255) / 255, ((c >> 8) & 255) / 255, (c & 255) / 255


class Writer:
    def __init__(self, fonts_dir):
        self.files = {"semibold": glob.glob(fonts_dir + "/*SegoeUI-Semibold*")[0],
                      "regular": [f for f in glob.glob(fonts_dir + "/*SegoeUI.*")][0]}
        self.fonts = {k: pymupdf.Font(fontfile=v) for k, v in self.files.items()}

    def kind(self, span):
        return "semibold" if "Semibold" in span["font"] else "regular"

    def put(self, page, x, y, text, kind, size, color):
        name = "XYC" + kind
        page.insert_font(fontname=name, fontfile=self.files[kind])
        page.insert_text((x, y), text, fontname=name, fontsize=size, color=rgb(color))

    def width(self, text, kind, size):
        return self.fonts[kind].text_length(text, size)


def erase(page, spans_):
    for s in spans_:
        page.add_redact_annot(pymupdf.Rect(s["bbox"]) + (0.2, 0.8, -0.2, -0.8), fill=False)
    page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE, graphics=pymupdf.PDF_REDACT_LINE_ART_NONE)


def replace_same_positions(page, w, span, new):
    """Lay out ``new`` from the span's first glyph, with the span's own letter spacing."""
    import statistics
    chars = span["chars"]
    k, size = w.kind(span), span["size"]
    gaps = [chars[i + 1]["origin"][0] - chars[i]["origin"][0] - w.width(chars[i]["c"], k, size)
            for i in range(len(chars) - 1)]
    spacing = statistics.median(gaps) if gaps else 0.0
    x, y = chars[0]["origin"]
    out = []
    for ch in new:
        out.append((x, y, ch, span))
        x += w.width(ch, k, size) + spacing
    return out


def main(orig, newpdf, pos, fonts_dir, out):
    doc = pymupdf.open(orig)
    src = pymupdf.open(newpdf)
    w = Writer(fonts_dir)
    n0 = doc.page_count

    # Portfolio page references (page 4).
    p4 = doc[3]
    edits, gone = [], []
    for s, t in list(spans(p4)):
        if t.upper().startswith("PAGE ") and "–" in t:
            a, b = (int(v) for v in t[5:].split("–"))
            if a <= pos - 1 <= b:
                b += 1
            elif a >= pos:
                a, b = a + 1, b + 1
            new = f"{t[:5]}{a:02d}–{b:02d}"
            if new != t:
                gone.append(s)
                edits += replace_same_positions(p4, w, s, new)
    erase(p4, gone)
    for x, y, ch, s in edits:
        w.put(p4, x, y, ch, w.kind(s), s["size"], s["color"])

    # Section openers: shift numbers, add the new line on its own opener.
    opener_for_new = None
    for i in range(n0):
        page = doc[i]
        all_spans = list(spans(page))
        if not any(t == "IN THIS SECTION" for _, t in all_spans):
            continue
        nums = [(s, t) for s, t in all_spans if t.strip().isdigit() and s["bbox"][1] < FOOTER_Y
                and 10 < s["size"] < 20]
        if nums and max(int(t) for _, t in nums) == pos - 1:
            opener_for_new = (i, nums, all_spans)
        edits, gone = [], []
        for s, t in nums:
            if int(t) >= pos:
                gone.append(s)
                edits += replace_same_positions(page, w, s, f"{int(t) + 1:02d}")
        if gone:
            erase(page, gone)
            for x, y, ch, s in edits:
                w.put(page, x, y, ch, w.kind(s), s["size"], s["color"])
    if opener_for_new:
        i, nums, all_spans = opener_for_new
        page = doc[i]
        last = max(nums, key=lambda st: int(st[1]))[0]
        ys = sorted({round(s["origin"][1], 1) for s, _ in nums})
        step = ys[1] - ys[0]
        title = min((s for s, t in all_spans if abs(s["origin"][1] - last["origin"][1]) < 0.5
                     and s["bbox"][0] > last["bbox"][2]), key=lambda s: s["bbox"][0])
        y = last["origin"][1] + step
        w.put(page, last["origin"][0], y, f"{pos:02d}", w.kind(last), last["size"], last["color"])
        w.put(page, title["origin"][0], y, NEW_TITLE, w.kind(title), title["size"], title["color"])

    # Printed page numbers after the insert (right-aligned).
    for i in range(pos - 1, n0):
        page = doc[i]
        s = next(s for s, t in spans(page) if s["bbox"][1] > FOOTER_Y and t.strip() == str(i + 1))
        right, y = s["bbox"][2], s["origin"][1]
        erase(page, [s])
        new = str(i + 2)
        k = w.kind(s)
        w.put(page, right - w.width(new, k, s["size"]), y, new, k, s["size"], s["color"])

    # The new page: <new-deck.pdf> may be the whole deck or a one-page export.
    src_page = pos - 1 if src.page_count > 1 else 0
    rect = doc[0].rect
    page = doc.new_page(pos - 1, width=rect.width, height=rect.height)
    page.show_pdf_page(page.rect, src, src_page)

    toc = doc.get_toc()
    if toc:
        doc.set_toc(toc)
    doc.save(out, garbage=4, deflate=True)
    print("saved", out, doc.page_count, "pages")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], int(sys.argv[3]), sys.argv[4], sys.argv[5])
