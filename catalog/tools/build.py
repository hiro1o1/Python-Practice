"""Build clean, English-labelled catalog images from the supplier board photos.

    pip install -r catalog/tools/requirements.txt
    python catalog/tools/build.py            # all boards
    python catalog/tools/build.py rk3588-a   # one board

Outputs, per board, in ``catalog/images/``:

* ``<id>.png`` / ``<id>.webp`` – labelled catalog image (PNG for print, WebP for web)
* ``<id>-plain.png`` – the cleaned board on white, no labels
* ``<id>-cutout.png`` – the cleaned board on a transparent background (for the PPTX catalog)

plus ``cutouts.json`` (crop origin and scale of each cutout, used to place port
markers from ``boards.py`` coordinates),

and regenerates ``catalog/index.html`` from ``boards.py``.
"""

import json
import os
import sys

import cv2
import numpy as np
from PIL import Image

sys.path.insert(0, os.path.dirname(__file__))
from boards import BOARDS  # noqa: E402
from clean import clean_board  # noqa: E402
from page import write_page  # noqa: E402
from render import render  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def plain(img, keep, margin=60):
    ys, xs = np.nonzero(keep)
    y0, y1 = max(0, ys.min() - margin), min(img.shape[0], ys.max() + margin)
    x0, x1 = max(0, xs.min() - margin), min(img.shape[1], xs.max() + margin)
    return Image.fromarray(img[y0:y1, x0:x1])


def cutout(img, keep, max_w=1000):
    """Tight crop with alpha; returns (image, crop origin, scale)."""
    ys, xs = np.nonzero(keep)
    x0, y0, x1, y1 = xs.min(), ys.min(), xs.max() + 1, ys.max() + 1
    alpha = cv2.GaussianBlur(keep[y0:y1, x0:x1].astype(np.float32) * 255, (3, 3), 0)
    im = Image.fromarray(np.dstack([img[y0:y1, x0:x1], alpha.astype(np.uint8)]), "RGBA")
    scale = min(1.0, max_w / im.width)
    if scale < 1:
        im = im.resize((round(im.width * scale), round(im.height * scale)), Image.LANCZOS)
    return im, [int(x0), int(y0)], scale


def main(ids):
    os.makedirs(os.path.join(ROOT, "images"), exist_ok=True)
    manifest_path = os.path.join(ROOT, "images", "cutouts.json")
    manifest = json.load(open(manifest_path)) if os.path.exists(manifest_path) else {}
    for cfg in BOARDS:
        if ids and cfg["id"] not in ids:
            continue
        src = np.asarray(Image.open(os.path.join(ROOT, "source", cfg["source"])).convert("RGB"))
        img, keep, _ = clean_board(src, cfg)
        out = os.path.join(ROOT, "images", cfg["id"])
        plain(img, keep).save(out + "-plain.png", optimize=True)
        labelled = render(img, keep, cfg)
        labelled.save(out + ".png", optimize=True)
        labelled.save(out + ".webp", quality=90, method=6)
        cut, origin, scale = cutout(img, keep)
        cut.save(out + "-cutout.png", optimize=True)
        manifest[cfg["id"]] = {"origin": origin, "scale": scale, "size": list(cut.size)}
        print("wrote", os.path.relpath(out + ".png", ROOT))
    with open(manifest_path, "w") as f:
        json.dump(manifest, f, indent=2)
    write_page(os.path.join(ROOT, "index.html"), BOARDS)
    print("wrote index.html")


if __name__ == "__main__":
    main(sys.argv[1:])
