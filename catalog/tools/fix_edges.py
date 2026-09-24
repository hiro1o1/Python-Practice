"""Targeted image fixes for the catalog PPTX; nothing else in the file is touched.

    python catalog/tools/fix_edges.py <in.pptx> <out.pptx>

1. White fringe on cut-out product images: the semi-transparent edge pixels
   still carry the white background they were cut from. Their colour is
   replaced with the colour of the nearest solid product pixels (alpha is
   kept), so no light line shows on dark or tinted pages.
2. Page 15 K-shaped kiosk stands: the three drawings sit on an off-white
   box that shows against the white cards. The box is made transparent and
   each drawing is scaled into the same frame with a small margin, so the
   small stand's base no longer reaches the card edge.

Only the affected media files inside the package are replaced; slide XML,
layout and all other images stay byte-identical.
"""

import io
import sys
import zipfile

import cv2
import numpy as np
from PIL import Image
from pptx import Presentation

SKIP_PREFIX = ("Glow", "Icon", "CE mark", "RoHS mark", "ISO mark")
KIOSK_PAGE = 15
KIOSK_NAMES = ("Small K-shaped", "Medium K-shaped", "Large K-shaped")


def decontaminate(rgba):
    a = rgba[..., 3]
    rgb = rgba[..., :3].astype(np.float32)
    solid = (a >= 250).astype(np.float32)
    todo = (a > 0) & (a < 250)
    out = rgb.copy()
    for k in (2, 4, 8, 16, 32):
        if not todo.any():
            break
        num = cv2.GaussianBlur(rgb * solid[..., None], (0, 0), k)
        den = cv2.GaussianBlur(solid, (0, 0), k)
        ok = todo & (den > 0.05)
        out[ok] = num[ok] / den[ok][:, None]
        todo &= ~ok
    return np.dstack([np.clip(out, 0, 255).astype(np.uint8), a])


def fringe_gap(rgba):
    a = rgba[..., 3].astype(int)
    lum = rgba[..., :3].astype(int).mean(-1)
    solid = (a >= 250).astype(np.uint8)
    inner = cv2.erode(solid, np.ones((3, 3), np.uint8)) - cv2.erode(solid, np.ones((9, 9), np.uint8))
    semi = (a > 0) & (a < 250)
    if not semi.any() or not inner.any():
        return 0
    return lum[semi].mean() - lum[inner.astype(bool)].mean()


def kiosk_cutout(rgb, frame_size, margin=0.07):
    x = rgb.astype(np.int32)
    w, sat = x.min(-1), x.max(-1) - x.min(-1)
    cand = ((w >= 248) & (sat < 8)).astype(np.uint8)
    n, lab = cv2.connectedComponents(cand)
    border = set(np.unique(np.r_[lab[0], lab[-1], lab[:, 0], lab[:, -1]])) - {0}
    bg = np.isin(lab, list(border))
    alpha = np.where(bg, 0, 255).astype(np.uint8)
    alpha = cv2.GaussianBlur(alpha, (3, 3), 0)
    alpha[~bg] = 255
    # Drop stray marks that are not part of the drawing (tiny separate specks).
    n, lab, stats, _ = cv2.connectedComponentsWithStats((alpha > 10).astype(np.uint8))
    big = stats[1:, 4].max()
    for j in range(1, n):
        if stats[j, 4] < 0.01 * big:
            alpha[lab == j] = 0
    im = decontaminate(np.dstack([rgb, alpha]))
    ys, xs = np.nonzero(im[..., 3] > 10)
    crop = Image.fromarray(im[ys.min():ys.max() + 1, xs.min():xs.max() + 1], "RGBA")
    fw, fh = frame_size
    s = min(fw * (1 - 2 * margin) / crop.width, fh * (1 - 2 * margin) / crop.height)
    crop = crop.resize((round(crop.width * s), round(crop.height * s)), Image.LANCZOS)
    canvas = Image.new("RGBA", (fw, fh), (255, 255, 255, 0))
    canvas.paste(crop, ((fw - crop.width) // 2, fh - round(fh * margin) - crop.height), crop)
    return canvas


def main(src, dst):
    prs = Presentation(src)
    replace = {}
    for i, slide in enumerate(prs.slides, 1):
        for sh in slide.shapes:
            if sh.shape_type != 13 or sh.name.startswith(SKIP_PREFIX):
                continue
            part = slide.part.related_part(sh._element.blipFill.blip.rEmbed)
            name = str(part.partname).lstrip("/")
            if name in replace:
                continue
            im = Image.open(io.BytesIO(part.blob))
            if i == KIOSK_PAGE and sh.name.startswith(KIOSK_NAMES):
                # 4x the original pixels so the drawing stays sharp at its size.
                big = im.convert("RGB").resize((im.width * 3, im.height * 3), Image.LANCZOS)
                out = kiosk_cutout(np.asarray(big), big.size)
                replace[name] = (out, f"page {i}: {sh.name}")
            elif im.mode == "RGBA":
                arr = np.asarray(im)
                if arr[..., 3].min() < 200 and fringe_gap(arr) > 25:
                    replace[name] = (Image.fromarray(decontaminate(arr), "RGBA"), f"page {i}: {sh.name}")
    zin = zipfile.ZipFile(src)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in replace:
                buf = io.BytesIO()
                replace[item.filename][0].save(buf, "PNG", optimize=True)
                data = buf.getvalue()
            zout.writestr(item, data)
    for k, (_, where) in sorted(replace.items()):
        print("fixed", k, "-", where)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
