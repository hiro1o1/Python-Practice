"""Stage 1: remove supplier annotations, watermarks and branding from a board photo."""

import os
import urllib.request

import cv2
import numpy as np

LAMA_URL = "https://github.com/Sanster/models/releases/download/add_big_lama/big-lama.pt"
LAMA_PATH = os.path.expanduser("~/.cache/catalog-tools/big-lama.pt")
_lama = None


def fill_holes(mask):
    ff = mask.copy()
    h, w = mask.shape
    cv2.floodFill(ff, np.zeros((h + 2, w + 2), np.uint8), (0, 0), 1)
    return mask | (1 - ff)


def remove_watermark(img, alpha, color):
    """Undo a semi-transparent colored text overlay.

    The overlay is ``out = (1 - a) * orig + a * color``.  Stroke pixels are found
    by their excess redness, the blend is inverted with a softened alpha, and the
    stroke area's chroma is re-filled from its surroundings to hide JPEG fringing.
    """
    a = img.astype(np.float32)
    r, g, b = a[..., 0], a[..., 1], a[..., 2]
    core = ((r - np.maximum(g, b) > 35) & (g - b < 20)).astype(np.uint8)
    core = cv2.morphologyEx(core, cv2.MORPH_OPEN, np.ones((2, 2), np.uint8))
    core = cv2.morphologyEx(core, cv2.MORPH_CLOSE, np.ones((3, 3), np.uint8))
    soft = np.clip(cv2.GaussianBlur(core.astype(np.float32), (0, 0), 1.0) * 1.15, 0, 1)
    al = (soft * alpha)[..., None]
    un = np.clip((a - al * np.array(color, np.float32)) / (1 - al), 0, 255).astype(np.uint8)
    ycc = cv2.cvtColor(un, cv2.COLOR_RGB2YCrCb)
    zone = cv2.dilate(core, cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (7, 7)))
    chroma = cv2.inpaint(np.dstack([ycc[..., 1], ycc[..., 2], ycc[..., 1]]), zone, 4, cv2.INPAINT_TELEA)
    ycc[..., 1], ycc[..., 2] = chroma[..., 0], chroma[..., 1]
    # Faint residue (anti-aliased stroke edges): pull any chroma that is redder than
    # its neighbourhood back to the local median.
    cr = ycc[..., 1]
    med = cv2.medianBlur(cr, 21)
    resid = (cr.astype(np.int16) - med > 5).astype(np.uint8)
    resid = cv2.dilate(resid, np.ones((3, 3), np.uint8))
    cr[resid > 0] = med[resid > 0]
    ycc[..., 1] = cr
    return cv2.cvtColor(ycc, cv2.COLOR_YCrCb2RGB)


def largest_component(mask):
    n, lab, stats, _ = cv2.connectedComponentsWithStats(mask, connectivity=8)
    if n <= 1:
        return mask
    return (lab == 1 + np.argmax(stats[1:, 4])).astype(np.uint8)


def board_mask(img, white=238):
    """Pixels belonging to the board and its connectors (not the white backdrop)."""
    nonbg = (img.min(-1) < white).astype(np.uint8)
    opened = cv2.morphologyEx(nonbg, cv2.MORPH_OPEN, np.ones((9, 9), np.uint8))
    n, lab, stats, _ = cv2.connectedComponentsWithStats(opened)
    core = fill_holes((lab == 1 + np.argmax(stats[1:, 4])).astype(np.uint8))
    return largest_component(fill_holes(nonbg & cv2.dilate(core, np.ones((31, 31), np.uint8))))


def annotation_mask(img, colors, protect):
    """Red leader lines/dots and yellow label boxes drawn by the supplier."""
    a = img.astype(np.int32)
    r, g, b = a[..., 0], a[..., 1], a[..., 2]
    mask = np.zeros(img.shape[:2], np.uint8)
    if "red" in colors:
        red = ((r - g > 70) & (r - b > 60) & (np.abs(g - b) < 45) & (r > 140)).astype(np.uint8)
        n, lab, stats, cent = cv2.connectedComponentsWithStats(red)
        for i in range(1, n):
            cx, cy = cent[i]
            if any(abs(cx - px) < 30 and abs(cy - py) < 30 for px, py in protect):
                continue
            mask[lab == i] = 1
    if "yellow" in colors:
        yellow = ((r > 150) & (g > 130) & (b < 140) & (r - b > 70) & (g - b > 60)
                  & (np.abs(r - g) < 70)).astype(np.uint8)
        # Solid label boxes survive close-then-open; thin outlines and leaders do not.
        solid = cv2.morphologyEx(cv2.morphologyEx(yellow | mask, cv2.MORPH_CLOSE, np.ones((9, 9), np.uint8)),
                                 cv2.MORPH_OPEN, np.ones((13, 13), np.uint8))
        n, lab, stats, _ = cv2.connectedComponentsWithStats(solid)
        for i in range(1, n):
            x, y, w, h, area = stats[i]
            mask[max(0, y - 6):y + h + 6, max(0, x - 6):x + w + 6] = 1
        n, lab, stats, _ = cv2.connectedComponentsWithStats(yellow)
        for i in range(1, n):
            if stats[i, 4] >= 40:
                mask[lab == i] = 1
    return cv2.dilate(mask, np.ones((7, 7), np.uint8))


def region_mask(img, regions):
    mask = np.zeros(img.shape[:2], np.uint8)
    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY).astype(np.int32)
    for kind, (x0, y0, x1, y1) in regions:
        sub = gray[y0:y1, x0:x1]
        if kind == "rect":
            m = np.ones_like(sub, np.uint8)
        elif kind == "light":
            m = (sub > np.percentile(sub, 50) + 45).astype(np.uint8)
            m = cv2.dilate(m, np.ones((5, 5), np.uint8))
        elif kind == "white":
            continue
        elif kind == "dark":
            m = (sub < np.percentile(sub, 50) - 25).astype(np.uint8)
            m = cv2.dilate(m, np.ones((5, 5), np.uint8))
        else:
            raise ValueError(kind)
        mask[y0:y1, x0:x1] |= m
    return mask


def arrow_mask(img, arrows):
    """Thin dark supplier arrows near each tip (only where they show against lighter parts)."""
    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
    blackhat = cv2.morphologyEx(gray, cv2.MORPH_BLACKHAT, np.ones((9, 9), np.uint8))
    thin_dark = ((blackhat > 45) & (gray < 70)).astype(np.uint8)
    band = np.zeros(gray.shape, np.uint8)
    for tip, center in arrows:
        cv2.line(band, tuple(tip), tuple(center), 1, 41)
        cv2.circle(band, tuple(tip), 18, 1, -1)
    return cv2.dilate(thin_dark & band, np.ones((5, 5), np.uint8))


def _load_lama():
    global _lama
    if _lama is None:
        import torch

        if not os.path.exists(LAMA_PATH):
            os.makedirs(os.path.dirname(LAMA_PATH), exist_ok=True)
            print("downloading LaMa weights ...")
            urllib.request.urlretrieve(LAMA_URL, LAMA_PATH)
        _lama = torch.jit.load(LAMA_PATH, map_location="cpu").eval()
    return _lama


def inpaint(img, mask, pad=160):
    """LaMa inpainting, run per cluster of masked pixels with surrounding context."""
    import torch

    model = _load_lama()
    out = img.copy()
    h, w = mask.shape
    groups = cv2.dilate(mask, np.ones((pad, pad), np.uint8))
    n, lab, stats, _ = cv2.connectedComponentsWithStats(groups)
    for i in range(1, n):
        x, y, bw, bh = stats[i, :4]
        x0, y0 = max(0, x), max(0, y)
        x1, y1 = min(w, x + bw), min(h, y + bh)
        x1 = x0 + max(8, (x1 - x0) // 8 * 8)
        y1 = y0 + max(8, (y1 - y0) // 8 * 8)
        x1, y1 = min(w, x1), min(h, y1)
        m = mask[y0:y1, x0:x1] * (lab[y0:y1, x0:x1] == i)
        if not m.any():
            continue
        crop = out[y0:y1, x0:x1]
        # Large holes fill more naturally at half resolution (LaMa's training scale).
        hole = cv2.boundingRect(m)
        scale = 0.5 if min(hole[2], hole[3]) > 40 else 1.0
        src = cv2.resize(crop, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA) if scale != 1 else crop
        msk = (cv2.resize(m, (src.shape[1], src.shape[0]), interpolation=cv2.INTER_NEAREST) if scale != 1 else m)
        if scale != 1:
            msk = cv2.dilate(msk, np.ones((3, 3), np.uint8))
        ch, cw = src.shape[:2]
        ph, pw = (-ch) % 8, (-cw) % 8
        crop_p = np.pad(src, ((0, ph), (0, pw), (0, 0)), mode="reflect")
        m_p = np.pad(msk, ((0, ph), (0, pw)))
        t_img = torch.from_numpy(np.ascontiguousarray(crop_p)).permute(2, 0, 1)[None].float() / 255
        t_m = torch.from_numpy(m_p)[None, None].float()
        with torch.no_grad():
            res = model(t_img, t_m)[0].permute(1, 2, 0).numpy()
        res = np.clip(res * 255, 0, 255).astype(np.uint8)[:ch, :cw]
        if scale != 1:
            res = cv2.resize(res, (crop.shape[1], crop.shape[0]), interpolation=cv2.INTER_CUBIC)
        crop[m > 0] = res[m > 0]
    return out


def clean_board(img, cfg):
    """Return (clean_rgb, board_mask, inpainted_mask)."""
    if cfg.get("watermark"):
        img = remove_watermark(img, *cfg["watermark"])
    img = img.copy()
    for _, center in cfg.get("arrows", []):  # numbered callout bubbles sit on the backdrop
        cv2.ellipse(img, tuple(center), (50, 44), 0, 0, 360, (255, 255, 255), -1)
    keep = board_mask(img)
    img[keep == 0] = 255

    for kind, (x0, y0, x1, y1) in cfg.get("erase", []):
        if kind == "white":  # stray backdrop marks that touch the board outline
            keep[y0:y1, x0:x1] = 0
            img[y0:y1, x0:x1] = 255

    mask = np.zeros(keep.shape, np.uint8)
    strokes = np.zeros(keep.shape, np.uint8)
    if cfg.get("annotation_colors"):
        strokes = annotation_mask(img, cfg["annotation_colors"], cfg.get("protect", []))
        mask |= strokes
    if cfg.get("arrows"):
        strokes |= arrow_mask(img, cfg["arrows"])
        mask |= strokes
    mask |= region_mask(img, cfg.get("erase", []))
    mask &= cv2.dilate(keep, np.ones((3, 3), np.uint8))

    # Thin strokes (leader lines, arrow shafts) are filled from their immediate
    # neighbours; only thick areas (label boxes, dots, stickers) go to LaMa.
    thick = cv2.morphologyEx(mask, cv2.MORPH_OPEN, np.ones((13, 13), np.uint8))
    thick = cv2.dilate(thick, np.ones((5, 5), np.uint8)) & mask
    img = cv2.inpaint(img, mask, 4, cv2.INPAINT_TELEA)
    img = inpaint(img, thick)
    # Inpainted stubs of outside leader lines that came out near-white at the
    # board edge are backdrop, not board.
    core = cv2.morphologyEx(keep, cv2.MORPH_OPEN, np.ones((9, 9), np.uint8))
    stub = strokes & (1 - core) & (img.min(-1) >= 238)
    keep = largest_component(keep & (1 - stub))
    img[keep == 0] = 255
    return img, keep, mask
