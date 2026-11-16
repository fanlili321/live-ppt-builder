"""直播 PPT 独立复核：Wind MCP 取不复权日K（每条线一次调用），与工具箱序列逐日比对，并核四点与最大回撤。
用法：python3 wind_spotcheck.py lines.json   （在 output/NNN_claude输出 下运行；Wind 结果缓存在 data/audit/wind_*.json）"""
import json, os, subprocess, glob

W = os.path.expanduser('~/.claude/skills/wind-mcp-skill')
A = 'data/audit'

import sys
# 配置文件：[[代码, 起, 止, 用途], ...]；工具箱序列从 data/audit/src_<代码>_*.json 读（final_audit 类脚本取数时缓存）
LINES = json.load(open(sys.argv[1])) if len(sys.argv) > 1 else []


def wind(code, a, b):
    fn = f'{A}/wind_{code}_{a}_{b}.json'
    if not os.path.exists(fn):
        arg = json.dumps({"windcode": code, "begin_date": a, "end_date": b, "aftype": "2"})
        r = subprocess.run(['node', 'scripts/cli.mjs', 'call', 'index_data', 'get_index_kline', arg],
                           cwd=W, capture_output=True, text=True, timeout=300)
        env = json.loads(r.stdout)
        if env.get('isError'):
            raise SystemExit(f'Wind 出错 {code}: {r.stdout[:400]}')
        open(fn, 'w').write(r.stdout)
    env = json.load(open(fn))
    d = json.loads(env['content'][0]['text'])['data']
    cols = [c['name'] for c in d['columns']]
    i, j = cols.index('TIME'), cols.index('MATCH')
    return {row[i][:10]: float(row[j]) for row in d['rows'] if row[j] not in (None, '')}


def toolbox(code, a, b):
    m = {}
    for fn in glob.glob(f'{A}/src_{code}_*.json'):
        m.update(json.load(open(fn)))
    return {k: v for k, v in m.items() if a <= k <= b}


def mdd(s):
    ks = sorted(s); pk, pv = ks[0], s[ks[0]]; best = (0, '', '')
    for k in ks:
        if s[k] > pv: pv, pk = s[k], k
        dd = s[k] / pv - 1
        if dd < best[0]: best = (dd, pk, k)
    return best


out = []
for code, a, b, use in LINES:
    w = wind(code, a, b)
    t = toolbox(code, a, b)
    common = sorted(set(w) & set(t))
    diffs = [(abs(w[k] / t[k] - 1), k) for k in common]
    worst = max(diffs) if diffs else (None, None)
    bad = [k for e, k in diffs if e > 0.0005]
    only_w = sorted(set(w) - set(t)); only_t = sorted(set(t) - set(w))
    ks = sorted(w)
    hi = max(ks, key=lambda k: w[k]); lo = min(ks, key=lambda k: w[k])
    mw, mt = mdd(w), mdd(t) if t else (None,)
    line = {
        'code': code, 'use': use, 'wind_rows': len(w), 'toolbox_rows': len(t), 'common': len(common),
        'max_rel_diff': worst[0], 'at': worst[1], 'n_bad(>0.05%)': len(bad), 'bad_sample': bad[:5],
        'only_wind': only_w[:5], 'only_toolbox': only_t[:5],
        'start': (ks[0], w[ks[0]], t.get(ks[0])), 'end': (ks[-1], w[ks[-1]], t.get(ks[-1])),
        'high': (hi, w[hi], t.get(hi)), 'low': (lo, w[lo], t.get(lo)),
        'cum_wind': w[ks[-1]] / w[ks[0]] - 1,
        'mdd_wind': mw, 'mdd_toolbox': mt,
    }
    out.append(line)
    print(f"{code:18s} {use}\n   Wind {len(w)} 行 / 工具箱 {len(t)} 行，重叠 {len(common)}，最大相对差 {worst[0]:.6f}@{worst[1]}，超 0.05% 的 {len(bad)} 天 {bad[:3]}"
          f"{'，仅Wind有 ' + str(only_w[:3]) if only_w else ''}{'，仅工具箱有 ' + str(only_t[:3]) if only_t else ''}\n"
          f"   四点 Wind/工具箱：起 {line['start']} 末 {line['end']} 高 {line['high']} 低 {line['low']}\n"
          f"   累计 {line['cum_wind']:.2%}；最大回撤 Wind {mw[0]:.2%} {mw[1]}→{mw[2]}，工具箱 {mt[0]:.2%} {mt[1]}→{mt[2]}")
json.dump(out, open(f'{A}/wind_spotcheck.json', 'w'), ensure_ascii=False, indent=1)
