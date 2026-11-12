#!/usr/bin/env python3
"""把 .key 里走势图外框、内容卡片框等「外框」形状的描边统一成标准：#FEAD00、6pt、实线（阴影状态不动）。

判定外框：形状描边为实线、粗细 ≥ 2pt、面积够大（宽 ≥ 400 且高 ≥ 100）。气泡的 1pt 细描边不算。
做法：直接改这些形状所用样式里的描边颜色/粗细（float 原地替换，长度不变）；样式是共用的，所以一次改全。
描边是从父样式继承来的（样式自身没写描边）时不改，会打印出来，需在 Keynote 里手动改一个后再跑。

用法：python3 set_frame_strokes.py 输入.key 输出.key [--dry-run]
之后：keynote_tool.py verify 输入 输出 → 打开导出页图目检。
"""
import sys, os, struct
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'keynote_iwa'))
import keynote_tool as kt
from pb import fields, rv, objects

STD_RGB = (0.995, 0.6806, 0.0)   # Keynote 色板橙黄（第 3 行第 4 格）
STD_W = 6.0


def objs(b):
    for lp, ais, aie, msgs in objects(b):
        oid = None
        for fn, wt, ts, vs, ve in fields(b, ais, aie):
            if fn == 1: oid, _ = rv(b, vs)
        for (ts, vs, ve), ps, pe in msgs:
            t = None
            for f2, w2, t2, v2s, v2e in fields(b, vs, ve):
                if f2 == 1: t, _ = rv(b, v2s)
            yield oid, t, ps, pe


def sub(b, s, e, fn):
    for f in fields(b, s, e):
        if f[0] == fn: return f


def main():
    a = [x for x in sys.argv[1:] if not x.startswith('--')]
    dry = '--dry-run' in sys.argv
    src, out = a[0], a[1]
    E = kt.read_entries(src)
    st = bytearray(kt.load_iwa_bytes(E['Index/DocumentStylesheet.iwa']))
    STY = {oid: (ps, pe) for oid, t, ps, pe in objs(bytes(st))}

    def own_stroke(sid):
        ps, pe = STY[sid]
        f1 = sub(st, ps, pe, 1)
        if not f1: return None, None
        f11 = sub(st, f1[3], f1[4], 1); parent = None
        if f11:
            f3 = sub(st, f11[3], f11[4], 3)
            if f3: parent = rv(st, sub(st, f3[3], f3[4], 1)[3])[0]
        p = sub(st, f1[3], f1[4], 11)
        sk = sub(st, p[3], p[4], 2) if p else None
        return sk, parent

    def eff_stroke(sid, d=0):
        if sid not in STY or d > 6: return None, None
        sk, parent = own_stroke(sid)
        if sk: return sk, sid
        return eff_stroke(parent, d + 1) if parent else (None, None)

    targets, skipped = {}, []
    for k in sorted(E):
        if not (k.startswith('Index/Slide') and k.endswith('.iwa')): continue
        b = kt.load_iwa_bytes(E[k])
        for oid, t, ps, pe in objs(b):
            if t != 2011: continue
            f1 = sub(b, ps, pe, 1); f12 = sub(b, f1[3], f1[4], 2)
            sid = rv(b, sub(b, f12[3], f12[4], 1)[3])[0]
            g = sub(b, sub(b, f1[3], f1[4], 1)[3], sub(b, f1[3], f1[4], 1)[4], 1)
            size = sub(b, g[3], g[4], 2)
            w, h = (struct.unpack('<f', b[x[3]:x[4]])[0] for x in fields(b, size[3], size[4]))
            sk, owner = eff_stroke(sid)
            if not sk: continue
            pat = sub(st, sk[3], sk[4], 6)
            if not pat or rv(st, sub(st, pat[3], pat[4], 1)[3])[0] != 1: continue
            wf = sub(st, sk[3], sk[4], 2)
            sw = struct.unpack('<f', st[wf[3]:wf[4]])[0] if wf else 0
            if sw < 2 or w < 400 or h < 100: continue
            if owner != sid and owner is not None:
                targets.setdefault(owner, []).append((k, oid))
            else:
                targets.setdefault(sid, []).append((k, oid))
    for sid, uses in targets.items():
        sk, _ = own_stroke(sid)
        col = sub(st, sk[3], sk[4], 1); wf = sub(st, sk[3], sk[4], 2)
        fs = [sub(st, col[3], col[4], i) for i in (3, 4, 5)] if col else []
        if not col or not wf or any(f is None or f[1] != 5 for f in fs) or wf[1] != 5:
            skipped.append((sid, uses)); continue
        for f, v in zip(fs, STD_RGB): st[f[3]:f[4]] = struct.pack('<f', v)
        st[wf[3]:wf[4]] = struct.pack('<f', STD_W)
        print('样式', sid, '用于', len(uses), '个外框', uses[:3])
    for s in skipped: print('跳过（描边字段不完整，需手动）', s)
    if dry or not targets: return
    kt.zip_update(src, out, {'Index/DocumentStylesheet.iwa': kt.dump_iwa_bytes(bytes(st))})
    print('→', out)


if __name__ == '__main__':
    main()
