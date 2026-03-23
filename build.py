#!/usr/bin/env python3
"""
build.py – generate icons and package TabBar.oxt

Run from the TabBar/ directory:
    python build.py

Outputs ../TabBar.oxt
"""

import os
import struct
import zlib
import zipfile

# ---------------------------------------------------------------------------
# Minimal PNG generator (no Pillow required)
# ---------------------------------------------------------------------------

def _chunk(tag: bytes, data: bytes) -> bytes:
    crc = zlib.crc32(tag + data) & 0xFFFFFFFF
    return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", crc)


def make_png(width: int, height: int, pixels) -> bytes:
    """
    pixels: list of (R, G, B) tuples, row-major order.
    Returns raw PNG bytes.
    """
    raw = b""
    for y in range(height):
        raw += b"\x00"  # filter = None
        for x in range(width):
            r, g, b = pixels[y * width + x]
            raw += bytes([r, g, b])
    sig  = b"\x89PNG\r\n\x1a\n"
    ihdr = _chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
    idat = _chunk(b"IDAT", zlib.compress(raw, 9))
    iend = _chunk(b"IEND", b"")
    return sig + ihdr + idat + iend


def draw_tabbar_icon(size: int) -> bytes:
    """
    Draw a simple tab bar icon: three tabs above a document body.
    The leftmost tab is active (white, merges into the body).
    The other two are inactive (grey).
    """
    W = H = size
    BG   = (230, 230, 230)   # background
    BODY = (255, 255, 255)   # document body / active tab
    TAB  = (180, 180, 200)   # inactive tab
    BORD = ( 80,  80,  80)   # border

    pixels = [BG] * (W * H)

    def px(x, y, c):
        if 0 <= x < W and 0 <= y < H:
            pixels[y * W + x] = c

    def rect(x0, y0, x1, y1, c, filled=True):
        for y in range(y0, y1 + 1):
            for x in range(x0, x1 + 1):
                if filled or x == x0 or x == x1 or y == y0 or y == y1:
                    px(x, y, c)

    s = size / 16.0
    def sc(v): return int(round(v * s))

    # Document body
    rect(sc(1), sc(5), sc(14), sc(14), BODY)
    rect(sc(1), sc(5), sc(14), sc(14), BORD, filled=False)

    # Active tab (leftmost) — white, merges into document body
    rect(sc(1), sc(2), sc(5), sc(5), BODY)
    rect(sc(1), sc(2), sc(5), sc(5), BORD, filled=False)
    # Remove the bottom border of the active tab so it flows into the document
    for x in range(sc(1) + 1, sc(5)):
        px(x, sc(5), BODY)

    # Inactive tab 2
    rect(sc(6), sc(3), sc(10), sc(5), TAB)
    rect(sc(6), sc(3), sc(10), sc(5), BORD, filled=False)

    # Inactive tab 3
    rect(sc(11), sc(3), sc(14), sc(5), TAB)
    rect(sc(11), sc(3), sc(14), sc(5), BORD, filled=False)

    return make_png(W, H, pixels)


# ---------------------------------------------------------------------------
# Build
# ---------------------------------------------------------------------------

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_OXT    = os.path.join(SCRIPT_DIR, "..", "TabBar.oxt")

ICONS_DIR = os.path.join(SCRIPT_DIR, "icons")
os.makedirs(ICONS_DIR, exist_ok=True)

print("Generating icons…")
for size, name in [(16, "tabbar_16.png"), (26, "tabbar_26.png")]:
    path = os.path.join(ICONS_DIR, name)
    with open(path, "wb") as f:
        f.write(draw_tabbar_icon(size))
    print(f"  {path}")

print("Packaging OXT…")
with zipfile.ZipFile(OUT_OXT, "w", zipfile.ZIP_DEFLATED) as z:
    for root, dirs, files in os.walk(SCRIPT_DIR):
        dirs[:] = [d for d in dirs if d not in ("__pycache__", "screenshots")]
        for fname in files:
            if fname == "build.py" or fname.endswith(".pyc"):
                continue
            full_path = os.path.join(root, fname)
            arc_name  = os.path.relpath(full_path, SCRIPT_DIR).replace("\\", "/")
            z.write(full_path, arc_name)
            print(f"  + {arc_name}")

print(f"\nDone -> {os.path.abspath(OUT_OXT)}")
print("Install via: Tools > Extension Manager > Add…")
