"""Stage 2: lay out English callouts around a cleaned board photo."""

import os

import numpy as np
from PIL import Image, ImageDraw, ImageFont

FONT_DIR = os.path.join(os.path.dirname(__file__), "fonts")
INK = (17, 24, 39)          # label text
MUTED = (100, 110, 125)     # sub-label text
LINE = (55, 65, 81)         # leader lines
ACCENT = (37, 99, 235)      # anchor dots


def font(weight, size):
    return ImageFont.truetype(os.path.join(FONT_DIR, f"Inter-{weight}.ttf"), size)


class Style:
    def __init__(self, s=1.0):
        self.title = font("SemiBold", round(23 * s))
        self.sub = font("Regular", round(18 * s))
        self.tag_title = font("SemiBold", round(17 * s))
        self.tag_sub = font("Regular", round(14 * s))
        self.line_w = max(2, round(2 * s))
        self.dot_r = round(7 * s)
        self.gap = round(10 * s)        # between stacked labels
        self.pad = round(14 * s)        # leader-line clearance from the board
        self.margin = round(60 * s)     # outer whitespace


def text_size(draw, text, f):
    if not text:
        return 0, 0
    l, t, r, b = draw.textbbox((0, 0), text, font=f)
    return r - l, b - t


def label_block(draw, c, st, tag=False):
    tf, sf = (st.tag_title, st.tag_sub) if tag else (st.title, st.sub)
    tw, th = text_size(draw, c["title"], tf)
    sw, sh = text_size(draw, c["sub"], sf)
    asc_t = tf.getmetrics()[0]
    line_gap = round(asc_t * 0.25)
    h = asc_t + (line_gap + sf.getmetrics()[0] if c["sub"] else 0)
    return max(tw, sw), h, asc_t, line_gap


def draw_label(draw, x, y, c, st, align, tag=False):
    """Draw at (x, y) = top-left for align='l', top-right for 'r', top-center for 'c'."""
    tf, sf = (st.tag_title, st.tag_sub) if tag else (st.title, st.sub)
    w, h, asc_t, gap = label_block(draw, c, st, tag)

    def put(text, f, yy, color):
        tw, _ = text_size(draw, text, f)
        xx = {"l": x, "r": x - tw, "c": x - tw / 2}[align]
        draw.text((xx, yy), text, font=f, fill=color, anchor="ls")

    put(c["title"], tf, y + asc_t, INK)
    if c["sub"]:
        put(c["sub"], sf, y + asc_t + gap + sf.getmetrics()[0], MUTED)


def spread(items, lo, hi):
    """1-D non-overlap: items are [want_start, size]; returns starts within [lo, hi]."""
    if not items:
        return []
    order = sorted(range(len(items)), key=lambda i: items[i][0])
    pos = [0.0] * len(items)
    cur = lo
    for i in order:
        pos[i] = max(items[i][0], cur)
        cur = pos[i] + items[i][1]
    # Shift the packed group back toward where its labels want to be, within bounds.
    shift = sum(items[i][0] - pos[i] for i in order) / len(order)
    first, last = pos[order[0]], pos[order[-1]] + items[order[-1]][1]
    shift = min(max(shift, lo - first), hi - last)
    if first + shift < lo:
        shift = lo - first
    pos = [p + shift for p in pos]
    return pos


def anchors(c):
    return c["at"] if isinstance(c["at"], list) else [c["at"]]


def mid(pts):
    return (sum(p[0] for p in pts) / len(pts), sum(p[1] for p in pts) / len(pts))


def render(board_img, keep, cfg, target_h=1300):
    """Return the final catalog image (RGB)."""
    ys, xs = np.nonzero(keep)
    bx0, by0, bx1, by1 = xs.min(), ys.min(), xs.max() + 1, ys.max() + 1
    s = float(np.clip(target_h / (by1 - by0), 1.0, 2.0))
    st = Style(max(1.0, s * 0.95))

    crop = Image.fromarray(board_img[by0:by1, bx0:bx1])
    bw, bh = round(crop.width * s), round(crop.height * s)
    crop = crop.resize((bw, bh), Image.LANCZOS)

    def tr(p):
        return ((p[0] - bx0) * s, (p[1] - by0) * s)

    scratch = ImageDraw.Draw(Image.new("RGB", (8, 8)))
    sides = {k: [c for c in cfg["callouts"] if c["side"] == k] for k in ("left", "right", "top", "bottom", "tag")}

    def widest(cs):
        return max([label_block(scratch, c, st)[0] for c in cs] or [0])

    def row_h(cs):
        return max([label_block(scratch, c, st)[1] for c in cs] or [0])

    # Top/bottom labels sit in one row, spread apart with angled leaders; a second
    # row is used only when the labels cannot fit across the canvas at all.
    def rows_for(cs):
        cs = sorted(cs, key=lambda c: mid([tr(p) for p in anchors(c)])[0])
        total = sum(label_block(scratch, c, st)[0] + st.gap * 2 for c in cs)
        room = bw + (widest(sides["left"]) + lead * 2 if sides["left"] else 0) + (widest(sides["right"]) + lead * 2 if sides["right"] else 0) + st.margin
        return cs, (1 if total <= room else 2)

    lead = round(40 * st.dot_r / 7)
    top_cs, top_rows = rows_for(sides["top"])
    bot_cs, bot_rows = rows_for(sides["bottom"])
    rh = row_h(sides["top"] + sides["bottom"]) + st.gap
    left_w = widest(sides["left"]) + lead * 2 if sides["left"] else 0
    right_w = widest(sides["right"]) + lead * 2 if sides["right"] else 0
    top_h = (rh * top_rows + lead) if top_cs else 0
    bot_h = (rh * bot_rows + lead) if bot_cs else 0

    ox, oy = st.margin + left_w, st.margin + top_h
    W, H = ox + bw + right_w + st.margin, oy + bh + bot_h + st.margin
    canvas = Image.new("RGB", (W, H), "white")
    canvas.paste(crop, (ox, oy))
    d = ImageDraw.Draw(canvas)

    def P(p):
        x, y = tr(p)
        return (x + ox, y + oy)

    lines, dots, labels = [], [], []

    # ---- left / right --------------------------------------------------------
    for side in ("left", "right"):
        cs = sorted(sides[side], key=lambda c: mid([P(p) for p in anchors(c)])[1])
        if not cs:
            continue
        blocks = [label_block(scratch, c, st) for c in cs]
        want = [mid([P(p) for p in anchors(c)])[1] - b[1] / 2 for c, b in zip(cs, blocks)]
        ys_ = spread([[w, b[1] + st.gap] for w, b in zip(want, blocks)], st.margin * 0.5, H - st.margin * 0.5)
        edge = ox - st.pad if side == "left" else ox + bw + st.pad
        text_x = ox - lead - st.pad if side == "left" else ox + bw + lead + st.pad
        elbow = ox - lead * 0.45 - st.pad if side == "left" else ox + bw + lead * 0.45 + st.pad
        for c, b, y in zip(cs, blocks, ys_):
            ly = y + b[2] * 0.62
            pts = [P(p) for p in anchors(c)]
            for p in pts:
                dots.append(p)
                lines.append([p, (edge, p[1]), (elbow, ly), (text_x - st.gap * (1 if side == "right" else -1), ly)])
            labels.append((text_x, y, c, "l" if side == "right" else "r", False))

    # ---- top / bottom --------------------------------------------------------
    for side, cs, nrows in (("top", top_cs, top_rows), ("bottom", bot_cs, bot_rows)):
        if not cs:
            continue
        blocks = [label_block(scratch, c, st) for c in cs]
        rows = [[] for _ in range(nrows)]
        for i, c in enumerate(cs):
            rows[i % nrows].append(i)
        for r, idx in enumerate(rows):
            items = []
            for i in idx:
                cx = mid([P(p) for p in anchors(cs[i])])[0]
                items.append([cx - blocks[i][0] / 2 - st.gap, blocks[i][0] + st.gap * 2])
            xs_ = spread(items, st.margin * 0.5, W - st.margin * 0.5)
            for i, x in zip(idx, xs_):
                c, b = cs[i], blocks[i]
                cx = x + st.gap + b[0] / 2
                if side == "top":
                    y = oy - lead - st.pad - rh * (r + 1) + st.gap
                    ly_end = y + b[1] + st.gap * 0.6
                    edge = oy - st.pad
                else:
                    y = oy + bh + st.pad + lead + rh * r
                    ly_end = y - st.gap * 0.6
                    edge = oy + bh + st.pad
                for p in [P(q) for q in anchors(c)]:
                    dots.append(p)
                    elbow = edge - lead * 0.45 if side == "top" else edge + lead * 0.45
                    lines.append([p, (p[0], edge), (cx, elbow), (cx, ly_end)])
                labels.append((cx, y, c, "c", False))

    # ---- on-board tags -------------------------------------------------------
    tags = []
    for c in sides["tag"]:
        w, h, _, _ = label_block(scratch, c, st, tag=True)
        tx, ty = P(c["tag"])
        px, py = round(9 * st.dot_r / 7), round(6 * st.dot_r / 7)
        box = (tx - w / 2 - px, ty - h / 2 - py, tx + w / 2 + px, ty + h / 2 + py)
        for p in [P(q) for q in anchors(c)]:
            dots.append(p)
            lines.append([p, (min(max(p[0], box[0]), box[2]), min(max(p[1], box[1]), box[3]))])
        tags.append((box, tx, ty - h / 2, c))

    for pts in lines:
        d.line(pts, fill="white", width=st.line_w + 4, joint="curve")
    for pts in lines:
        d.line(pts, fill=LINE, width=st.line_w, joint="curve")
    for box, tx, ty, c in tags:
        d.rounded_rectangle(box, radius=round(6 * st.dot_r / 7), fill=(255, 255, 255), outline=LINE,
                            width=max(1, st.line_w - 1))
        draw_label(d, tx, ty, c, st, "c", tag=True)
    r = st.dot_r
    for x, y in dots:
        d.ellipse((x - r - 2, y - r - 2, x + r + 2, y + r + 2), fill="white")
        d.ellipse((x - r, y - r, x + r, y + r), fill=ACCENT)
    for x, y, c, align, tag in labels:
        draw_label(d, x, y, c, st, align, tag)
    return canvas
