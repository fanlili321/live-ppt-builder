"""Minimal IWA (Keynote) chunk reader/writer with pure-python snappy."""
import struct


def _varint(b, i):
    r = s = 0
    while True:
        c = b[i]; i += 1
        r |= (c & 0x7F) << s; s += 7
        if c < 0x80:
            return r, i


def snappy_decompress(b):
    n, i = _varint(b, 0)
    out = bytearray()
    while i < len(b):
        tag = b[i]; i += 1
        t = tag & 3
        if t == 0:
            ln = tag >> 2
            if ln >= 60:
                k = ln - 59
                ln = int.from_bytes(b[i:i + k], 'little'); i += k
            ln += 1
            out += b[i:i + ln]; i += ln
            continue
        if t == 1:
            ln = ((tag >> 2) & 7) + 4
            off = ((tag >> 5) << 8) | b[i]; i += 1
        elif t == 2:
            ln = (tag >> 2) + 1
            off = int.from_bytes(b[i:i + 2], 'little'); i += 2
        else:
            ln = (tag >> 2) + 1
            off = int.from_bytes(b[i:i + 4], 'little'); i += 4
        for _ in range(ln):
            out.append(out[-off])
    assert len(out) == n, (len(out), n)
    return bytes(out)


def snappy_compress_literal(b):
    out = bytearray()
    n = len(b)
    while True:
        out.append(n & 0x7F | (0x80 if n > 0x7F else 0)); n >>= 7
        if not n:
            break
    i = 0
    while i < len(b):
        chunk = b[i:i + 65536]
        ln = len(chunk) - 1
        if ln < 60:
            out.append(ln << 2)
        elif ln < 256:
            out += bytes([60 << 2, ln])
        else:
            out += bytes([61 << 2]) + ln.to_bytes(2, 'little')
        out += chunk
        i += len(chunk)
    return bytes(out)


def read_iwa(path):
    b = open(path, 'rb').read()
    i = 0
    out = bytearray()
    while i < len(b):
        assert b[i] == 0, b[i]
        ln = int.from_bytes(b[i + 1:i + 4], 'little')
        out += snappy_decompress(b[i + 4:i + 4 + ln])
        i += 4 + ln
    return bytes(out)


def write_iwa(path, data):
    out = bytearray()
    for i in range(0, len(data), 65536):
        c = snappy_compress_literal(data[i:i + 65536])
        out += b'\x00' + len(c).to_bytes(3, 'little') + c
    open(path, 'wb').write(bytes(out))
