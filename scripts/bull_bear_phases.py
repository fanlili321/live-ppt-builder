"""按 20% 规则划分牛熊市（486 期王磊反馈 P14 用，见 rules/数据口径与常用代码.md「牛熊市划分」）。
- 从高点跌 20% 确认熊市，从低点涨 20% 确认牛市；阶段起止 = 峰/谷日（收盘价）。
- 默认把熊市中的反弹并入熊市：熊-牛-熊 中间的「牛」没创新高、后一段熊又跌破前低 → 合并为一轮熊市（用户 2026-11-13 定）。
- 区间开头到第一个峰/谷的不足 20% 的小段标为 lead-in，不算一轮。
用法: python3 bull_bear_phases.py 日频.json [--out phases.json] [--no-merge]
  日频.json = {"YYYY-MM-DD": 收盘, ...}（或 [[日期, 收盘], ...]）"""
import json, sys, argparse
ap = argparse.ArgumentParser(); ap.add_argument('src'); ap.add_argument('--out'); ap.add_argument('--no-merge', action='store_true')
ap.add_argument('--th', type=float, default=0.2, help='阈值，默认 0.2')
a = ap.parse_args()
d = json.load(open(a.src))
items = d.items() if isinstance(d, dict) else d
s = sorted((k, float(v)) for k, v in items if v not in (None, ''))
th = a.th
# 1) 原始 20% 拐点
pts = [0]; trend = None; ext = 0
for i, (_, v) in enumerate(s):
    if trend is None:
        hi = max(range(pts[-1], i + 1), key=lambda j: s[j][1]); lo = min(range(pts[-1], i + 1), key=lambda j: s[j][1])
        if v <= s[hi][1] * (1 - th): trend = 'down'; pts.append(hi); ext = i
        elif v >= s[lo][1] * (1 + th): trend = 'up'; pts.append(lo); ext = i
        continue
    if trend == 'up':
        if v > s[ext][1]: ext = i
        elif v <= s[ext][1] * (1 - th): pts.append(ext); trend = 'down'; ext = i
    else:
        if v < s[ext][1]: ext = i
        elif v >= s[ext][1] * (1 + th): pts.append(ext); trend = 'up'; ext = i
if trend: pts.append(len(s) - 1 if ext != len(s) - 1 else ext)
ph = []
for x, y in zip(pts, pts[1:]):
    if y <= x: continue
    ph.append(['bull' if s[y][1] > s[x][1] else 'bear', x, y])
# 2) 合并熊市中的反弹
if not a.no_merge:
    k = 0
    while k + 2 < len(ph):
        b1, r, b2 = ph[k], ph[k + 1], ph[k + 2]
        if b1[0] == r[0] == b2[0] or not (b1[0] == 'bear' and r[0] == 'bull' and b2[0] == 'bear'): k += 1; continue
        if s[r[2]][1] < s[b1[1]][1] and s[b2[2]][1] < s[b1[2]][1]:
            ph[k:k + 3] = [['bear', b1[1], b2[2]]]; continue
        k += 1
# 3) 开头不足阈值的小段
out = []
for i, (kd, x, y) in enumerate(ph):
    r = s[y][1] / s[x][1] - 1
    if i == 0 and x == 0 and abs(r) < th: kd = 'lead-in'
    out.append(dict(kind=kd, start=s[x][0], end=s[y][0], v0=s[x][1], v1=s[y][1], ret=r))
    print(f'{kd:7s} {s[x][0]} → {s[y][0]}  {s[x][1]:.2f} → {s[y][1]:.2f}  {r*100:+.2f}%')
if a.out: json.dump(dict(phases=out), open(a.out, 'w'), ensure_ascii=False, indent=1)
