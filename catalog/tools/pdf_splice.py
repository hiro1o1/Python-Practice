"""Insert the new catalog page into the PowerPoint-exported PDF.

    python catalog/tools/pdf_splice.py <original.pdf> <new-deck.pdf> <page> <font.ttf> <out.pdf>

The original pages are kept exactly as exported. The only edits are the
printed page numbers after the insert and the fixed "Page NN–NN" references
on page 4, which are swapped digit by digit in the same font, size, colour and
letter spacing. ``font.ttf`` is Segoe UI Semibold (e.g. the subset embedded in
the original PDF).
"""

import statistics
import sys

import pymupdf

PAGE_REFS = {"05–10": "05–11", "11–13": "12–14", "14–15": "15–16", "16–18": "17–19"}
FOOTER_Y = 800  # page numbers sit below this line (points)


def spans(page):
    for b in page.get_text("rawdict")["blocks"]:
        for line in b.get("lines", []):
            for s in line["spans"]:
                yield s, "".join(c["c"] for c in s["chars"])


def rgb(color):
    return ((color >> 16) & 255) / 255, ((color >> 8) & 255) / 255, (color & 255) / 255


def replace_chars(page, span, start, new, font, fname):
    """Replace span text from char index ``start`` with ``new``, keeping spacing."""
    chars = span["chars"]
    size = span["size"]
    gaps = [chars[i + 1]["origin"][0] - chars[i]["origin"][0] - font.text_length(chars[i]["c"], size)
            for i in range(len(chars) - 1) if chars[i]["c"] != " "]
    spacing = statistics.median(gaps) if gaps else 0.0
    x0, y = chars[start]["origin"]
    bbox = pymupdf.Rect(chars[start]["bbox"]) | chars[-1]["bbox"]
    page.add_redact_annot(bbox + (0.2, 0.6, -0.2, -0.6), fill=False)
    page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE,
                          graphics=pymupdf.PDF_REDACT_LINE_ART_NONE)
    return [(x, y, ch, size, span["color"]) for x, ch in _layout(x0, new, font, size, spacing)]


def _layout(x, text, font, size, spacing):
    for ch in text:
        yield x, ch
        x += font.text_length(ch, size) + spacing


def main(orig_path, new_path, page_no, font_path, out_path):
    doc = pymupdf.open(orig_path)
    src = pymupdf.open(new_path)
    font = pymupdf.Font(fontfile=font_path)
    fname = "SegoeSB"

    # 1. Fixed page references on page 4.
    p4 = doc[3]
    writes = []
    for s, t in list(spans(p4)):
        for old, new in PAGE_REFS.items():
            if t.upper().endswith(old):
                start = len(t) - len(old)
                writes += replace_chars(p4, s, start, new, font, fname)
    p4.insert_font(fontname=fname, fontbuffer=font.buffer)
    for x, y, ch, size, color in writes:
        p4.insert_text((x, y), ch, fontname=fname, fontsize=size, color=rgb(color))

    # 2. Printed page numbers on every page after the insert.
    for idx in range(page_no - 1, doc.page_count):
        page = doc[idx]
        for s, t in list(spans(page)):
            if s["bbox"][1] > FOOTER_Y and t.strip().isdigit() and int(t) == idx + 1:
                right, y = s["bbox"][2], s["origin"][1]
                page.add_redact_annot(pymupdf.Rect(s["bbox"]) + (0.2, 0.6, -0.2, -0.6), fill=False)
                page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE,
                                      graphics=pymupdf.PDF_REDACT_LINE_ART_NONE)
                new = str(idx + 2)
                page.insert_font(fontname=fname, fontbuffer=font.buffer)
                page.insert_text((right - font.text_length(new, s["size"]), y), new,
                                 fontname=fname, fontsize=s["size"], color=rgb(s["color"]))
                break
        else:
            raise SystemExit(f"page number not found on page {idx + 1}")

    # 3. The new page, drawn at the catalog's exact page size.
    w, h = doc[0].rect.width, doc[0].rect.height
    page = doc.new_page(page_no - 1, width=w, height=h)
    page.show_pdf_page(page.rect, src, page_no - 1)

    doc.save(out_path, garbage=3, deflate=True)
    print("saved", out_path, doc.page_count, "pages")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], int(sys.argv[3]), sys.argv[4], sys.argv[5])
