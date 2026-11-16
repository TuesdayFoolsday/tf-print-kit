"""QuickXorHash, the fingerprint OneDrive / SharePoint keep for every file (Graph: file.hashes.quickXorHash).

160-bit state; byte i is XORed in at bit (11 * i) mod 160, wrapping around; at the end the file length (64 bits,
little-endian) is XORed into the last 8 bytes; the 20 state bytes are written base64. Empty input -> "AAAAAAAAAAAAAAAAAAAAAAAAAAA=".
11 and 160 share no factor, so the bit positions repeat every 160 bytes: the hash keeps one XOR-folded 160-byte row
and works through the data in big chunks without a Python loop per byte (a 1 GB file takes seconds, not minutes).

    h = QuickXor(); h.update(chunk); ...; h.b64()          quickxor_file(path) -> (base64, size)
"""
from __future__ import annotations

import base64
from pathlib import Path

WIDTH = 160
ROW = 160          # bytes after which the bit positions repeat
MASK = (1 << WIDTH) - 1


class QuickXor:
    def __init__(self) -> None:
        self.n = 0
        self.row = 0   # int: byte j holds the XOR of all input bytes whose index % 160 == j

    def update(self, data: bytes) -> None:
        data = bytes(data)
        pos = self.n % ROW
        self.n += len(data)
        head = min(len(data), (ROW - pos) % ROW)   # bytes until the next row boundary
        for k in range(head):
            self.row ^= data[k] << (8 * (pos + k))
        rest = data[head:]
        rows = len(rest) // ROW
        if rows:
            self.row ^= _fold(int.from_bytes(rest[:rows * ROW], "little"), rows)
        for k, b in enumerate(rest[rows * ROW:]):
            self.row ^= b << (8 * k)

    def digest(self) -> bytes:
        state = 0
        for j in range(ROW):
            v = (self.row >> (8 * j)) & 0xFF
            if v:
                shifted = v << ((11 * j) % WIDTH)
                state ^= (shifted | (shifted >> WIDTH)) & MASK    # rotate within 160 bits
        state ^= (self.n & (2**64 - 1)) << (WIDTH - 64)
        return state.to_bytes(WIDTH // 8, "little")

    def b64(self) -> str:
        return base64.b64encode(self.digest()).decode()


def _fold(x: int, rows: int) -> int:
    """XOR of all 160-byte rows of x (little-endian int of `rows` rows)."""
    acc, bits = 0, ROW * 8
    while rows > 1:
        half = rows // 2
        part = (1 << (half * bits)) - 1
        if rows % 2:
            acc ^= x >> (2 * half * bits)
        x = (x & part) ^ ((x >> (half * bits)) & part)
        rows = half
    return acc ^ x


def quickxor_bytes(data: bytes) -> str:
    h = QuickXor()
    h.update(data)
    return h.b64()


def quickxor_file(path, chunk: int = 8 * 1024 * 1024) -> tuple[str, int]:
    """-> (base64 hash, size) reading the file in chunks."""
    h = QuickXor()
    with Path(path).open("rb") as f:
        while piece := f.read(chunk):
            h.update(piece)
    return h.b64(), h.n
