#!/usr/bin/env python3
"""Generate the PWA icon set with no third-party dependencies.

Draws a flat calendar mark on a deep navy field. Run from the repo root:
    python3 tools/make-icons.py
"""
import struct
import zlib
from pathlib import Path

BG = (0x24, 0x3B, 0x53)      # deep slate navy
INK = (0xFF, 0xFF, 0xFF)     # calendar body
TODAY = (0xF0, 0xA3, 0x3D)   # amber "today" marker

OUT = Path(__file__).resolve().parent.parent / "docs" / "icons"


class Canvas:
    def __init__(self, size, bg):
        self.size = size
        self.px = bytearray(bg * size * size)

    def rect(self, x, y, w, h, color):
        """Fill an axis-aligned rectangle, clipped to the canvas."""
        x0, y0 = max(0, int(x)), max(0, int(y))
        x1, y1 = min(self.size, int(x + w)), min(self.size, int(y + h))
        if x1 <= x0 or y1 <= y0:
            return
        row = bytes(color) * (x1 - x0)
        for yy in range(y0, y1):
            start = (yy * self.size + x0) * 3
            self.px[start:start + len(row)] = row

    def png(self, path):
        raw = bytearray()
        stride = self.size * 3
        for y in range(self.size):
            raw.append(0)  # filter type 0
            raw += self.px[y * stride:(y + 1) * stride]

        def chunk(tag, data):
            body = tag + data
            return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)

        png = b"\x89PNG\r\n\x1a\n"
        png += chunk(b"IHDR", struct.pack(">IIBBBBB", self.size, self.size, 8, 2, 0, 0, 0))
        png += chunk(b"IDAT", zlib.compress(bytes(raw), 9))
        png += chunk(b"IEND", b"")
        path.write_bytes(png)
        return path


def draw(size, mark_fraction):
    """Draw the calendar mark centred on the field.

    mark_fraction is the share of the canvas the mark occupies: keep it small
    for maskable icons so the mark survives an aggressive circular crop.
    """
    c = Canvas(size, BG)
    m = size * mark_fraction
    x0 = y0 = (size - m) / 2
    border = max(2.0, m * 0.075)

    # Frame: white block, then punch the interior back to the background.
    c.rect(x0, y0, m, m, INK)
    c.rect(x0 + border, y0 + border, m - 2 * border, m - 2 * border, BG)

    # Header bar across the top of the interior.
    head_h = m * 0.17
    c.rect(x0 + border, y0 + border, m - 2 * border, head_h, INK)

    # 3x2 grid of day cells below the header.
    body_top = y0 + border + head_h + m * 0.085
    body_left = x0 + border + m * 0.10
    body_w = m - 2 * border - 2 * (m * 0.10)
    cell = body_w / 3.9
    gap = (body_w - 3 * cell) / 2
    for row in range(2):
        for col in range(3):
            cx = body_left + col * (cell + gap)
            cy = body_top + row * (cell + gap)
            color = TODAY if (row, col) == (0, 1) else INK
            c.rect(cx, cy, cell, cell, color)
    return c


if __name__ == "__main__":
    OUT.mkdir(parents=True, exist_ok=True)
    targets = [
        ("icon-192.png", 192, 0.70),
        ("icon-512.png", 512, 0.70),
        ("icon-maskable-512.png", 512, 0.52),
        ("apple-touch-icon.png", 180, 0.70),
    ]
    for name, size, frac in targets:
        p = draw(size, frac).png(OUT / name)
        print(f"{p.name}  {size}x{size}  {p.stat().st_size} bytes")
