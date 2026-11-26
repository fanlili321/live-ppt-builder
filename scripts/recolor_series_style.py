#!/usr/bin/env python3
"""改图表某个系列样式（type 5028）的折线颜色：python3 recolor_series_style.py in.key out.key 样式id R,G,B
R,G,B 为 0–1 小数（如创业板50全收益深蓝灰 0.411,0.458,0.635）。改的是线条描边（field 10000/48）和图例色块（10000/172）。
样式 id 用 swap_pie_wedge_style.py ... --list 查（第 n 个=第 n 条线）。
490 期经验：同一页要换颜色，先用 set_series_style.py 把系列引用换成「同一张图表家族里」闲置的系列样式
（如该图第 3、4 个样式），再用本脚本给它上色；直接引用别的期次图表的样式 id 会渲染成别的颜色（实测变成蓝/绿）。
两页共用同一组样式时（同一页复制出来的），改样式颜色会同时影响这几页，改前先 --list 看清。
"""
import sys, os, struct
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'keynote_iwa'))
import keynote_tool as k, find_strokes as fs
src, dst, sid, rgb = sys.argv[1], sys.argv[2], int(sys.argv[3]), [float(x) for x in sys.argv[4].split(',')]
e = k.read_entries(src); ds = bytearray(k.load_iwa_bytes(e['Index/DocumentStylesheet.iwa']))
n = 0
for ident, t, path, w, c, off in fs.scan(bytes(ds)):
    if ident != sid or path not in ([10000, 48], [10000, 172]): continue
    i = off - 1
    while i > off - 80:
        if ds[i] == 0x1d and ds[i + 5] == 0x25 and ds[i + 10] == 0x2d:
            v = [struct.unpack('<f', ds[i + 1 + 5 * j:i + 5 + 5 * j])[0] for j in range(3)]
            if all(abs(a - b) < 0.006 for a, b in zip(v, c)): break
        i -= 1
    else:
        raise SystemExit(f'没找到颜色字段 {path}')
    for j, val in enumerate(rgb): ds[i + 1 + 5 * j:i + 5 + 5 * j] = struct.pack('<f', val)
    n += 1; print('改', path, c, '→', tuple(rgb))
k.zip_update(src, dst, {'Index/DocumentStylesheet.iwa': k.dump_iwa_bytes(bytes(ds))}); print(n, '处 →', dst)
