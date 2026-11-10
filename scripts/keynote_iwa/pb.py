"""Generic protobuf helpers for IWA streams."""


def rv(b, i):
    r = s = 0
    while True:
        c = b[i]; i += 1
        r |= (c & 0x7F) << s; s += 7
        if c < 0x80:
            return r, i


def ev(n):
    out = bytearray()
    while True:
        c = n & 0x7F; n >>= 7
        if n:
            out.append(c | 0x80)
        else:
            out.append(c); return bytes(out)


def fields(b, start, end):
    """Yield (field_no, wire_type, tag_start, value_start, value_end)."""
    i = start
    while i < end:
        ts = i
        tag, i = rv(b, i)
        fn, wt = tag >> 3, tag & 7
        if wt == 0:
            _, j = rv(b, i); yield fn, wt, ts, i, j; i = j
        elif wt == 1:
            yield fn, wt, ts, i, i + 8; i += 8
        elif wt == 5:
            yield fn, wt, ts, i, i + 4; i += 4
        elif wt == 2:
            ln, j = rv(b, i); yield fn, wt, ts, j, j + ln; i = j + ln
        else:
            raise ValueError("bad wire type %d at %d" % (wt, ts))
    if i != end:
        raise ValueError("overrun")


def try_fields(b, s, e):
    try:
        return list(fields(b, s, e))
    except Exception:
        return None


def objects(b):
    """Yield (archive_info_len_pos, ai_start, ai_end, [(msg_info_field_span, payload_start, payload_end)])."""
    i = 0
    while i < len(b):
        lp = i
        ln, i = rv(b, i)
        ais, aie = i, i + ln
        payload = aie
        msgs = []
        for fn, wt, ts, vs, ve in fields(b, ais, aie):
            if fn == 2:
                L = None
                for f2, w2, t2, v2s, v2e in fields(b, vs, ve):
                    if f2 == 3:
                        L, _ = rv(b, v2s)
                msgs.append(((ts, vs, ve), payload, payload + L))
                payload += L
        yield lp, ais, aie, msgs
        i = payload
