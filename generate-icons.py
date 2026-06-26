#!/usr/bin/env python3
"""Generates minimal solid-color PNG icons. No Pillow required — pure stdlib."""
import struct, zlib, os

def make_png(w, h, r, g, b):
    """Create a minimal solid-color PNG."""
    def chunk(tag, data):
        c = struct.pack('>I', len(data)) + tag + data
        return c + struct.pack('>I', zlib.crc32(tag + data) & 0xffffffff)
    ihdr = struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0)
    raw = b''.join(b'\x00' + bytes([r, g, b] * w) for _ in range(h))
    return (b'\x89PNG\r\n\x1a\n'
            + chunk(b'IHDR', ihdr)
            + chunk(b'IDAT', zlib.compress(raw))
            + chunk(b'IEND', b''))

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets')
os.makedirs(out, exist_ok=True)
for size, name in [(32, 'icon-32.png'), (80, 'icon-80.png')]:
    with open(os.path.join(out, name), 'wb') as f:
        f.write(make_png(size, size, 26, 115, 232))  # #1a73e8 blue
    print(f"Created assets/{name}")
