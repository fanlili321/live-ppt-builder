#!/usr/bin/env python3
"""自动把每个 ▶ 三角对准它右边那行文字的正中间（垂直方向）。

做法：Keynote 导出页图 → 逐页找 ▶（品牌橙小三角）和它右侧最近的一行文字，
量出「文字行中心 − 三角中心」→ 用 AppleScript 把对应的三角形状（44×39 那类小形状）移过去 →
再导出复查，直到所有 ▶ 偏差 ≤ 1px。不改文字、不改其他对象。

用法：
  python3 align_triangles.py "文档名.key"            # 文档须已在 Keynote 里打开（按名字引用）
  python3 align_triangles.py "文档名.key" --dry-run  # 只量不改
  --offset=N   三角中心比文字中心高 N px（默认 0 = 正中间）
  --slides=5,6,9  只处理这些页（幻灯片序号，含跳过页在内的序号）
"""
import sys, os, json, subprocess, tempfile, shutil
import numpy as np
from PIL import Image
from scipy import ndimage


def osa(script):
    r = subprocess.run(['osascript', '-'], input=script, capture_output=True, text=True, timeout=600)
    if r.returncode:
        raise SystemExit('AppleScript 出错：' + r.stderr.strip())
    return r.stdout.strip()


def export(doc, outdir):
    osa(f'''with timeout of 600 seconds
tell application "Keynote"
export document "{doc}" to (POSIX file "{outdir}") as slide images with properties {{image format:PNG}}
end tell
end timeout''')
    fs = sorted(f for f in os.listdir(outdir) if f.lower().endswith('.png'))
    return [os.path.join(outdir, f) for f in fs]


def measure(path, scale):
    """返回 [(三角左x, 三角上y, 三角中心y, 文字行中心y)]，单位：幻灯片坐标。"""
    a = np.asarray(Image.open(path).convert('RGB')).astype(int)
    r, g, b = a[..., 0], a[..., 1], a[..., 2]
    m = (r > 200) & (g > 90) & (g < 160) & (b < 90)
    lab, _ = ndimage.label(m)
    out = []
    for i, sl in enumerate(ndimage.find_objects(lab)):
        y0, y1, x0, x1 = sl[0].start, sl[0].stop, sl[1].start, sl[1].stop
        h, w = y1 - y0, x1 - x0
        if not (20 * scale <= h <= 60 * scale and 15 * scale <= w <= 60 * scale):
            continue
        fill = (lab[sl] == i + 1).sum() / (h * w)
        if not (0.35 <= fill <= 0.65):                   # 三角约占外框一半；圆点约 0.78，排除
            continue
        tc = (y0 + y1) / 2
        top, bot = max(0, y0 - int(40 * scale)), y1 + int(40 * scale)
        reg = a[top:bot, x1 + int(5 * scale): x1 + int(700 * scale)]
        dark = reg.sum(2) < 300
        rows = np.where(dark.sum(1) > 3)[0]
        if not len(rows):
            continue
        rows = rows + top
        bands, s, prev = [], rows[0], rows[0]
        for y in rows[1:]:
            if y > prev + 3:
                bands.append((s, prev)); s = y
            prev = y
        bands.append((s, prev))
        bd = min(bands, key=lambda bb: abs((bb[0] + bb[1]) / 2 - tc))
        lc = (bd[0] + bd[1]) / 2
        if abs(lc - tc) > 40 * scale:
            continue
        out.append((x0 / scale, y0 / scale, tc / scale, lc / scale))
    return out


def shapes_of(doc, n):
    s = osa(f'''tell application "Keynote"
set s to slide {n} of document "{doc}"
set out to ""
set k to 0
repeat with sh in shapes of s
 set k to k + 1
 set w to width of sh
 set h to height of sh
 if w < 80 and h < 80 then
  set p to position of sh
  set out to out & k & "," & (item 1 of p) & "," & (item 2 of p) & "," & w & "," & h & linefeed
 end if
end repeat
return out
end tell''')
    res = []
    for line in s.splitlines():
        if line.strip():
            k, x, y, w, h = [float(v) for v in line.split(',')]
            res.append((int(k), x, y, w, h))
    return res


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    doc = args[0]
    dry = '--dry-run' in sys.argv
    off = next((float(a.split('=')[1]) for a in sys.argv if a.startswith('--offset=')), 0.0)
    only = next((set(int(x) for x in a.split('=')[1].split(',')) for a in sys.argv if a.startswith('--slides=')), None)
    sw = float(osa(f'tell application "Keynote" to get width of document "{doc}"'))
    ignore = set()   # 附近没有可挪的小形状 = 图表里的橙色拐角等误识别，后续轮次忽略
    for rnd in range(3):
        tmp = tempfile.mkdtemp(prefix='tri_')
        try:
            pages = export(doc, tmp)
            scale = Image.open(pages[0]).size[0] / sw
            # 导出会略过「跳过」的页：导出第 i 张 = 第 vis[i-1] 张幻灯片
            sk = osa(f'tell application "Keynote" to get skipped of slides of document "{doc}"').split(', ')
            vis = [i for i, x in enumerate(sk, 1) if x.strip() == 'false']
            assert len(vis) == len(pages), ('导出页数与未跳过页数不符', len(vis), len(pages))
            todo = []
            for n, p in zip(vis, pages):
                if only and n not in only:
                    continue
                for tx, ty, tc, lc in measure(p, scale):
                    d = lc - off - tc
                    todo.append((n, tx, ty, tc, lc, d))
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
        todo = [t for t in todo if (t[0], round(t[1]), round(t[2])) not in ignore]
        bad = [t for t in todo if abs(t[5]) > 1]
        print(f'第{rnd + 1}轮：找到 ▶ {len(todo)} 个，需调整 {len(bad)} 个')
        for n, tx, ty, tc, lc, d in todo:
            print(f'  P{n} ▶ y={ty:.0f} 中心{tc:.1f} 文字中心{lc:.1f} 偏差{d:+.1f}')
        if dry or not bad:
            return
        moves = {}
        for n, tx, ty, tc, lc, d in bad:
            cands = [s for s in shapes_of(doc, n) if abs(s[1] - tx) < 12 and abs(s[2] - ty) < 12]
            if len(cands) != 1:
                if not cands:
                    ignore.add((n, round(tx), round(ty)))
                    print(f'  P{n} ({tx:.0f},{ty:.0f}) 附近没有形状，是图表等里的橙色块，忽略')
                else:
                    print(f'  P{n} 在 ({tx:.0f},{ty:.0f}) 附近找到 {len(cands)} 个形状，跳过，请人工检查')
                continue
            k, x, y, w, h = cands[0]
            moves.setdefault(n, []).append((k, x, round(y + d)))
        cmds = []
        for n, lst in moves.items():
            for k, x, y in lst:
                cmds.append(f'set position of shape {k} of slide {n} of d to {{{x:.0f}, {y}}}')
        if not cmds:
            continue
        osa('tell application "Keynote"\nset d to document "%s"\n%s\nsave d\nend tell' % (doc, '\n'.join(cmds)))
    print('3 轮后仍有偏差，请人工检查')


if __name__ == '__main__':
    main()
