#!/usr/bin/env python3
"""把 .key 里所有图片统一成标准边框（#FEAD00、6pt、实线、无阴影）。

做法：找到文档里已有的「标准边框」图片样式（stroke 颜色≈#FEAD00、宽 6、实线、无阴影），
把其他图片的样式引用换成它（同长度 varint 原地替换），并同步 Metadata 里该页的外部引用登记。
文档里还没有标准样式时：先在 Keynote 里手动给任意一张图设好（边框→线条→色板橙黄→6 点），保存后再跑。

用法：
  python3 set_image_borders.py 输入.key 输出.key [--dry-run] [--only=Slide-123.iwa,Slide-456.iwa]
  --only：只改这些页（页里有色带/卡片这类做装饰用的图片时，用它避开）
之后必须：keynote_tool.py verify 输入 输出（无缺图）→ Keynote 打开导出页图目检。
注意：图片自带（烤在像素里）的边框不会被去掉，需另外裁图（如二维码黄边约 18px）。
"""
import sys, os, struct
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'keynote_iwa'))
import keynote_tool as kt
from pb import fields, rv, objects, ev, try_fields


def objs(b):
    for lp, ais, aie, msgs in objects(b):
        oid = None
        for fn, wt, ts, vs, ve in fields(b, ais, aie):
            if fn == 1: oid, _ = rv(b, vs)
        for (ts, vs, ve), ps, pe in msgs:
            t = None
            for f2, w2, t2, v2s, v2e in fields(b, vs, ve):
                if f2 == 1: t, _ = rv(b, v2s)
            yield oid, t, ais, ps, pe


def sub(b, s, e, fn):
    for f in fields(b, s, e):
        if f[0] == fn: return f
    return None


def style_info(st, ps, pe):
    """返回 (parent, (r,g,b), width, pattern, shadow_on) ；没有自身描边返回 stroke=None"""
    parent = None
    f1 = sub(st, ps, pe, 1)
    if f1:
        f3 = sub(st, f1[3], f1[4], 3)
        if f3: parent = rv(st, sub(st, f3[3], f3[4], 1)[3])[0]
    props = sub(st, ps, pe, 11)
    if not props: return parent, None
    stroke = sub(st, props[3], props[4], 1)
    shadow = sub(st, props[3], props[4], 3)
    sh_on = False
    if shadow:
        f6 = sub(st, shadow[3], shadow[4], 6)
        sh_on = bool(f6 and rv(st, f6[3])[0])
    if not stroke: return parent, (None, None, None, sh_on)
    col = sub(st, stroke[3], stroke[4], 1)
    rgb = tuple(round(struct.unpack('<f', st[f[3]:f[4]])[0], 3) for k in (3, 4, 5) for f in [sub(st, col[3], col[4], k)] if f) if col else None
    w = sub(st, stroke[3], stroke[4], 2); w = struct.unpack('<f', st[w[3]:w[4]])[0] if w else None
    pat = sub(st, stroke[3], stroke[4], 6); pat = rv(st, sub(st, pat[3], pat[4], 1)[3])[0] if pat else None
    return parent, (rgb, w, pat, sh_on)


def main():
    a = [x for x in sys.argv[1:] if not x.startswith('--')]
    dry = '--dry-run' in sys.argv
    only = [x.split('=', 1)[1].split(',') for x in sys.argv[1:] if x.startswith('--only=')]
    only = ['Index/' + n for n in only[0]] if only else None
    src, out = a[0], a[1]
    E = kt.read_entries(src)
    st = kt.load_iwa_bytes(E['Index/DocumentStylesheet.iwa'])
    std = None
    for oid, t, ais, ps, pe in objs(st):
        if t != 3016: continue
        _, s = style_info(st, ps, pe)
        if s and s[0] and abs(s[0][0] - 0.995) < 0.02 and abs(s[0][1] - 0.681) < 0.02 and s[0][2] < 0.02 and s[1] == 6.0 and s[2] == 1 and not s[3]:
            std = oid; break
    if not std:
        sys.exit('文档里没有标准边框样式：先在 Keynote 里手动给一张图设好（#FEAD00/6点/实线/无阴影）再跑')
    print('标准样式 id', std)
    files, changed = {}, []
    meta = bytearray(kt.load_iwa_bytes(E['Index/Metadata.iwa']))
    for k in sorted(E):
        if not (k.startswith('Index/Slide') and k.endswith('.iwa')): continue
        if only and k not in only: continue
        b = bytearray(kt.load_iwa_bytes(E[k])); dirty = []
        for oid, t, ais, ps, pe in objs(bytes(b)):
            if t != 3005: continue
            f3 = sub(b, ps, pe, 3)
            sid = rv(b, sub(b, f3[3], f3[4], 1)[3])[0]
            if sid == std: continue
            old, new = ev(sid), ev(std)
            if len(old) != len(new):
                print('跳过（编号长度不同，需手动）', k, oid, sid); continue
            seg = bytes(b[ais:pe]); b[ais:pe] = seg.replace(old, new); dirty.append(sid); changed.append((k, oid, sid))
        if not dirty: continue
        nb = bytes(b); loc = k[len('Index/'):-len('.iwa')].encode()
        for oid, t, ais, ps, pe in objs(bytes(meta)):
            if t != 11006: continue
            for fn, wt, ts, vs, ve in fields(meta, ps, pe):
                if fn != 3: continue
                if loc not in [bytes(meta[x[3]:x[4]]) for x in fields(meta, vs, ve) if x[0] == 3]: continue
                seg = bytes(meta[vs:ve])
                if ev(std) in seg: continue
                for sid in set(dirty):
                    if ev(sid) in nb: continue
                    pat = b'\x10' + ev(sid)
                    if seg.count(pat) == 1:
                        seg = seg.replace(pat, b'\x10' + ev(std)); break
                meta[vs:ve] = seg
        files[k] = kt.dump_iwa_bytes(nb)
    for c in changed: print('改', *c)
    if dry or not changed: return
    files['Index/Metadata.iwa'] = kt.dump_iwa_bytes(bytes(meta))
    kt.zip_update(src, out, files)
    print('→', out)


if __name__ == '__main__':
    main()
