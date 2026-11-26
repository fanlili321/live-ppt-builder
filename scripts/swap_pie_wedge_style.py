#!/usr/bin/env python3
"""同一页两个饼图「从大到小排序 + 同一行业同色」：对调某一个饼图里两块扇形的样式引用。

为什么需要：Keynote 饼图的扇形颜色按行序号分配（第 n 行用第 n 个系列样式）。两个指数的行业大小顺序
不同时（490 期：创业板50 金融6.85%>医疗4.87%，创业板指 医疗7.83%>金融5.37%），都按从大到小排，
同一行业就会变成不同颜色。做法：两个图都按从大到小写数据（keynote_tool.py set），
再把其中一个图里那两块扇形的样式引用对调，颜色就跟另一个图对上了。

用法：
  python3 swap_pie_wedge_style.py in.key out.key Slide-xxx.iwa 第几个图(0起) 扇形a 扇形b [--list]
  扇形序号从 1 起（=数据第几行）。--list 只列出该页图表里的系列样式引用和位置，不改。
  例（490 期 P4，第 0 个图=创业板50，对调第 3、4 块）：
  python3 swap_pie_wedge_style.py in.key out.key Slide-8867739.iwa 0 3 4
原理：两个图共用同一组系列样式对象（同 id，各引用一次，按图表对象先后出现）。系列样式 id 在图表对象里
连续排列；第 n 块扇形 = 该图里第 n 个系列样式引用。只在指定那个图的对象范围内对调两处等长 varint。
改完：keynote_tool.py verify → Keynote 打开导出页图，目检两个饼图同一行业同色、顺序从大到小。
"""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'keynote_iwa'))
from keynote_tool import read_entries, load_iwa_bytes, dump_iwa_bytes, zip_update
from pb import fields, objects, rv, ev


def ids_of_type(b, t):
    out = []
    for lp, ais, aie, msgs in objects(b):
        oid = None; types = []
        for fn, wt, ts, vs, ve in fields(b, ais, aie):
            if fn == 1: oid = rv(b, vs)[0]
            if fn == 2:
                for f2, w2, t2, v2s, v2e in fields(b, vs, ve):
                    if f2 == 1: types.append(rv(b, v2s)[0])
        ps = msgs[0][1] if msgs else aie; pe = msgs[-1][2] if msgs else aie
        out.append((oid, types, ps, pe))
    return out


def main():
    a = sys.argv[1:]
    src, dst, sf, ci = a[0], a[1], a[2], int(a[3])
    E = read_entries(src); name = [n for n in E if n.endswith('/' + sf) or n == sf][0]
    b = bytearray(load_iwa_bytes(E[name]))
    charts = [o for o in ids_of_type(bytes(b), 0) if o[1] and o[1][0] == 5021]
    ss = load_iwa_bytes(E['Index/DocumentStylesheet.iwa'])
    series = sorted(o[0] for o in ids_of_type(ss, 0) if o[1] and o[1][0] == 5028)
    ch = charts[ci]
    refs = []
    for sid in series:
        pat = b'\x08' + ev(sid); i = ch[2]
        while True:
            i = bytes(b).find(pat, i, ch[3])
            if i < 0: break
            refs.append((i, sid)); i += 1
    refs.sort()
    # 连续等间距的一组 = 系列样式列表（扇形 1..n）
    groups = []; cur = []
    for r in refs:
        if cur and r[0] - cur[-1][0] not in (len(b'\x08' + ev(r[1])) + 6, 11, 12, 10):
            groups.append(cur); cur = []
        cur.append(r)
    if cur: groups.append(cur)
    lst = max(groups, key=len)
    if '--list' in a:
        print('图表对象', [c[0] for c in charts], '当前第', ci, '个')
        for k, (pos, sid) in enumerate(lst, 1): print(f'扇形{k}: 样式 {sid} @ {pos}')
        return
    wa, wb = int(a[4]), int(a[5])
    (pa, sa), (pb_, sb) = lst[wa - 1], lst[wb - 1]
    A, B = b'\x08' + ev(sa), b'\x08' + ev(sb)
    assert len(A) == len(B), 'varint 长度不同，不能原地对调'
    b[pa:pa + len(A)] = B; b[pb_:pb_ + len(B)] = A
    zip_update(src, dst, {name: dump_iwa_bytes(bytes(b))})
    print(f'对调 扇形{wa}({sa}) <-> 扇形{wb}({sb}) → {dst}')


if __name__ == '__main__':
    main()
