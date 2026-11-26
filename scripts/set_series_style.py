#!/usr/bin/env python3
"""让某页图表的第 n 条线（系列）直接用另一页同一指数那条线的样式（颜色、粗细完全一样）。

为什么需要：同一份 PPT 里同一个指数每页颜色要一致（rules/走势图配色与指数代码.md）。拼页后各页图表的
系列样式来自不同期次，颜色/粗细不一样。做法：把目标图表里第 n 个系列样式引用，换成参考页那条线的样式 id
（同为 type 5028，等长 varint 原地替换）。

用法：
  python3 set_series_style.py in.key out.key Slide-目标.iwa 图序号(0起) 系列n(1起)=样式id [系列m=样式id ...]
  先用 swap_pie_wedge_style.py ... --list 看参考页和目标页各系列的样式 id。
  例（490 期 P9，本图样式 8875781/82/83/84，82=浅蓝已是 P19 创业板指全收益的颜色）：
  python3 set_series_style.py in.key out.key Slide-8892427.iwa 0 1=8875782 2=8875783
  再 python3 recolor_series_style.py out.key out2.key 8875783 0.411,0.458,0.635  （创业板50全收益深蓝灰，同 P8）
⚠ 490 期实测：引用「别的期次图表」的样式 id 会渲染成别的颜色（变蓝/绿）。只在同一张图表家族的样式之间换引用（例如把第 1 条线换成本图第 2 个样式），要新颜色再用 recolor_series_style.py 上色。
注意：两页共用同一组系列样式时（同一张页复制出来的），改引用只影响目标页。改完 verify + 导出目检。
"""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'keynote_iwa'))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from keynote_tool import read_entries, load_iwa_bytes, dump_iwa_bytes, zip_update
from pb import ev
from swap_pie_wedge_style import ids_of_type


def series_refs(b, ch, series):
    refs = []
    for sid in series:
        pat = b'\x08' + ev(sid); i = ch[2]
        while True:
            i = b.find(pat, i, ch[3])
            if i < 0: break
            refs.append((i, sid)); i += 1
    refs.sort()
    groups, cur = [], []
    for r in refs:
        if cur and r[0] - cur[-1][0] not in (10, 11, 12):
            groups.append(cur); cur = []
        cur.append(r)
    if cur: groups.append(cur)
    return max(groups, key=len)


def main():
    src, dst, sf, ci = sys.argv[1], sys.argv[2], sys.argv[3], int(sys.argv[4])
    E = read_entries(src); name = [n for n in E if n.endswith('/' + sf)][0]
    b = bytearray(load_iwa_bytes(E[name]))
    ss = load_iwa_bytes(E['Index/DocumentStylesheet.iwa'])
    series = sorted(o[0] for o in ids_of_type(ss, 0) if o[1] and o[1][0] == 5028)
    ch = [o for o in ids_of_type(bytes(b), 0) if o[1] and o[1][0] == 5021][ci]
    lst = series_refs(bytes(b), ch, series)
    for spec in sys.argv[5:]:
        n, new = (int(x) for x in spec.split('='))
        pos, old = lst[n - 1]
        A, B = b'\x08' + ev(old), b'\x08' + ev(new)
        assert len(A) == len(B), 'varint 长度不同'
        assert new in series, f'{new} 不是系列样式'
        b[pos:pos + len(A)] = B
        print(f'系列{n}: {old} → {new}')
    zip_update(src, dst, {name: dump_iwa_bytes(bytes(b))}); print('→', dst)


if __name__ == '__main__':
    main()
