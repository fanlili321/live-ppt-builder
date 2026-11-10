#!/usr/bin/env python3
"""把底稿图表里的旧数据，和同一代码新取的数据，在重叠日期上逐个比对，判断代码/口径是否用对。

用法：
  python3 compare_history.py old.csv new.csv [--series 伦敦金=伦敦金,上海金=上海金] [--tol 0.005]

old.csv：`keynote_tool.py get` 导出的底稿图表数据（第一列日期，其余列=系列）。
new.csv：新取的数据，格式相同（第一列日期 YYYY-MM-DD，其余列=系列）。
--series：旧系列名=新系列名，默认按同名配对。
--tol：相对误差阈值，默认 0.5%。

输出每个系列：重叠日期数、吻合比例、最大误差及日期、结论（✅ 吻合 / ⚠️ 不吻合需问用户）。
有任一系列不吻合时退出码为 1。
"""
import argparse, csv, sys


def load(path):
    rows = list(csv.reader(open(path, encoding='utf-8-sig')))
    head = [h.strip() for h in rows[0]]
    data = {h: {} for h in head[1:]}
    for r in rows[1:]:
        if not r or not r[0].strip():
            continue
        d = r[0].strip()[:10]
        for h, v in zip(head[1:], r[1:]):
            if v.strip():
                data[h][d] = float(v)
    return data


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('old'); ap.add_argument('new')
    ap.add_argument('--series', default='')
    ap.add_argument('--tol', type=float, default=0.005)
    a = ap.parse_args()
    old, new = load(a.old), load(a.new)
    pairs = [p.split('=', 1) for p in a.series.split(',') if p] or [(s, s) for s in old if s in new]
    if not pairs:
        sys.exit(f'没有可配对的系列：旧={list(old)} 新={list(new)}，请用 --series 指定')
    bad = False
    for o, n in pairs:
        common = sorted(set(old.get(o, {})) & set(new.get(n, {})))
        if not common:
            print(f'⚠️ {o} ↔ {n}：没有重叠日期（检查频率：周频日期要对周五/月末）'); bad = True; continue
        errs = [(abs(new[n][d] / old[o][d] - 1) if old[o][d] else abs(new[n][d]), d) for d in common]
        ok = sum(1 for e, _ in errs if e <= a.tol) / len(errs)
        worst = max(errs)
        verdict = '✅ 吻合' if ok >= 0.98 else '⚠️ 不吻合，需问用户（代码/价格或全收益/频率/复权可能不同）'
        if ok < 0.98:
            bad = True
        print(f'{verdict}  {o} ↔ {n}：重叠 {len(common)} 个日期，{ok:.1%} 在 ±{a.tol:.1%} 内；'
              f'最大误差 {worst[0]:.2%}（{worst[1]}：旧 {old[o][worst[1]]} / 新 {new[n][worst[1]]}）')
    sys.exit(1 if bad else 0)


if __name__ == '__main__':
    main()
