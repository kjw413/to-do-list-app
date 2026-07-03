"""Generate the TodoList app icon.

The script intentionally uses only the Python standard library so the icon can
be regenerated on a clean machine before packaging the app.
"""

from __future__ import annotations

import math
import struct
import zlib
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "assets"
PNG_PATH = ASSETS / "todolist.png"
ICO_PATH = ASSETS / "todolist.ico"
SIZES = (16, 24, 32, 48, 64, 128, 256)


def rgba(hex_color: str, alpha: int = 255) -> tuple[int, int, int, int]:
    hex_color = hex_color.lstrip("#")
    return (
        int(hex_color[0:2], 16),
        int(hex_color[2:4], 16),
        int(hex_color[4:6], 16),
        alpha,
    )


class Canvas:
    def __init__(self, width: int, height: int):
        self.width = width
        self.height = height
        self.pixels = bytearray(width * height * 4)

    def _blend(self, x: int, y: int, color: tuple[int, int, int, int]) -> None:
        if x < 0 or y < 0 or x >= self.width or y >= self.height:
            return
        sr, sg, sb, sa = color
        if sa <= 0:
            return
        i = (y * self.width + x) * 4
        dr, dg, db, da = self.pixels[i : i + 4]
        inv = 255 - sa
        out_a = sa + (da * inv + 127) // 255
        if out_a <= 0:
            self.pixels[i : i + 4] = bytes((0, 0, 0, 0))
            return
        out_r = (sr * sa + dr * da * inv // 255) // out_a
        out_g = (sg * sa + dg * da * inv // 255) // out_a
        out_b = (sb * sa + db * da * inv // 255) // out_a
        self.pixels[i : i + 4] = bytes((out_r, out_g, out_b, out_a))

    def round_rect(
        self,
        x0: float,
        y0: float,
        x1: float,
        y1: float,
        radius: float,
        color: tuple[int, int, int, int],
    ) -> None:
        left = max(0, math.floor(x0))
        top = max(0, math.floor(y0))
        right = min(self.width, math.ceil(x1))
        bottom = min(self.height, math.ceil(y1))
        radius = max(0.0, radius)
        for y in range(top, bottom):
            py = y + 0.5
            for x in range(left, right):
                px = x + 0.5
                dx = max(x0 + radius - px, 0.0, px - (x1 - radius))
                dy = max(y0 + radius - py, 0.0, py - (y1 - radius))
                if dx * dx + dy * dy <= radius * radius:
                    self._blend(x, y, color)

    def circle(self, cx: float, cy: float, radius: float, color: tuple[int, int, int, int]) -> None:
        left = max(0, math.floor(cx - radius))
        top = max(0, math.floor(cy - radius))
        right = min(self.width, math.ceil(cx + radius))
        bottom = min(self.height, math.ceil(cy + radius))
        r2 = radius * radius
        for y in range(top, bottom):
            py = y + 0.5
            for x in range(left, right):
                px = x + 0.5
                if (px - cx) * (px - cx) + (py - cy) * (py - cy) <= r2:
                    self._blend(x, y, color)

    def line(
        self,
        x0: float,
        y0: float,
        x1: float,
        y1: float,
        width: float,
        color: tuple[int, int, int, int],
    ) -> None:
        half = width / 2.0
        left = max(0, math.floor(min(x0, x1) - half))
        top = max(0, math.floor(min(y0, y1) - half))
        right = min(self.width, math.ceil(max(x0, x1) + half))
        bottom = min(self.height, math.ceil(max(y0, y1) + half))
        vx = x1 - x0
        vy = y1 - y0
        length2 = vx * vx + vy * vy
        if length2 == 0:
            self.circle(x0, y0, half, color)
            return
        for y in range(top, bottom):
            py = y + 0.5
            for x in range(left, right):
                px = x + 0.5
                t = ((px - x0) * vx + (py - y0) * vy) / length2
                t = max(0.0, min(1.0, t))
                nx = x0 + t * vx
                ny = y0 + t * vy
                if (px - nx) * (px - nx) + (py - ny) * (py - ny) <= half * half:
                    self._blend(x, y, color)


def downsample(src: Canvas, size: int, scale: int) -> bytes:
    out = bytearray(size * size * 4)
    area = scale * scale
    for y in range(size):
        for x in range(size):
            r = g = b = a = 0
            for sy in range(scale):
                for sx in range(scale):
                    i = ((y * scale + sy) * src.width + (x * scale + sx)) * 4
                    r += src.pixels[i]
                    g += src.pixels[i + 1]
                    b += src.pixels[i + 2]
                    a += src.pixels[i + 3]
            j = (y * size + x) * 4
            out[j : j + 4] = bytes((r // area, g // area, b // area, a // area))
    return bytes(out)


def write_png_bytes(width: int, height: int, pixels: bytes) -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)

    scanlines = bytearray()
    stride = width * 4
    for y in range(height):
        scanlines.append(0)
        start = y * stride
        scanlines.extend(pixels[start : start + stride])

    return b"".join(
        (
            b"\x89PNG\r\n\x1a\n",
            chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 6, 0, 0, 0)),
            chunk(b"IDAT", zlib.compress(bytes(scanlines), 9)),
            chunk(b"IEND", b""),
        )
    )


def render_icon(size: int) -> bytes:
    scale = 4
    canvas = Canvas(size * scale, size * scale)
    k = canvas.width / 256.0

    def s(value: float) -> float:
        return value * k

    # App tile.
    canvas.round_rect(s(16), s(16), s(240), s(240), s(48), rgba("#2563eb"))
    canvas.round_rect(s(28), s(28), s(228), s(126), s(38), rgba("#38bdf8", 100))

    # Paper shadow and checklist sheet.
    canvas.round_rect(s(67), s(58), s(205), s(219), s(17), rgba("#0f172a", 58))
    canvas.round_rect(s(56), s(47), s(194), s(208), s(17), rgba("#fffaf0"))
    canvas.round_rect(s(66), s(64), s(184), s(82), s(5), rgba("#dbe7ff"))

    # Top clip.
    canvas.round_rect(s(92), s(32), s(158), s(70), s(12), rgba("#f59e0b"))
    canvas.round_rect(s(109), s(27), s(141), s(53), s(9), rgba("#cbd5e1"))
    canvas.round_rect(s(116), s(33), s(134), s(46), s(6), rgba("#f8fafc"))

    # Checklist rows.
    check = rgba("#16a34a")
    text = rgba("#334155")
    muted = rgba("#64748b")
    rows = (104, 139, 174)
    for index, y in enumerate(rows):
        canvas.round_rect(s(76), s(y - 11), s(98), s(y + 11), s(5), rgba("#e2e8f0"))
        canvas.line(s(80), s(y), s(87), s(y + 7), s(9), check)
        canvas.line(s(87), s(y + 7), s(101), s(y - 10), s(9), check)
        line_color = text if index == 0 else muted
        canvas.line(s(116), s(y - 6), s(169), s(y - 6), s(8), line_color)
        canvas.line(s(116), s(y + 8), s(154), s(y + 8), s(7), rgba("#94a3b8"))

    # Small folded corner for an unmistakable paper silhouette.
    canvas.line(s(171), s(47), s(194), s(70), s(4), rgba("#dbe7ff"))
    canvas.round_rect(s(176), s(50), s(191), s(72), s(4), rgba("#e0f2fe", 180))

    return downsample(canvas, size, scale)


def main() -> None:
    ASSETS.mkdir(exist_ok=True)
    png_entries = []
    for size in SIZES:
        pixels = render_icon(size)
        png_entries.append((size, write_png_bytes(size, size, pixels)))
        if size == 256:
            PNG_PATH.write_bytes(png_entries[-1][1])

    offset = 6 + 16 * len(png_entries)
    header = struct.pack("<HHH", 0, 1, len(png_entries))
    entries = bytearray()
    images = bytearray()
    for size, png in png_entries:
        width_byte = 0 if size >= 256 else size
        entries.extend(struct.pack("<BBBBHHII", width_byte, width_byte, 0, 0, 1, 32, len(png), offset))
        images.extend(png)
        offset += len(png)

    ICO_PATH.write_bytes(header + bytes(entries) + bytes(images))
    print(f"wrote {PNG_PATH}")
    print(f"wrote {ICO_PATH}")


if __name__ == "__main__":
    main()
