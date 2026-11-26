"""487 期终审：从工具箱重新取数，与在线文档图表数据逐点对照，并重算表格/标注数字。
用法：python3 scripts/final_audit.py   （在 output/487_claude输出 下运行）
"""
import csv, json, os, subprocess, datetime as dt

D = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'dbms_api.py')
A = 'data/audit'


def fetch(code, start, end):
    fn = f'{A}/src_{code}_{start}_{end}.json'
    if os.path.exists(fn):
        return json.load(open(fn))
    out = {}
    s = dt.date.fromisoformat(start)
    e = dt.date.fromisoformat(end)
    while s <= e:
        t = min(dt.date(s.year + 2, 12, 31), e)
        q = json.dumps({"codes": [code], "beginDate": str(s), "endDate": str(t),
                        "frequency": "daily", "fields": ["close"]})
        d = None
        for _ in range(3):
            r = subprocess.run(['python3', D, 'index-query', q], capture_output=True, text=True, timeout=170)
            try:
                d = json.loads(r.stdout)['data']
                break
            except Exception:
                d = None
        if d is None:
            raise SystemExit(f'取数失败 {code} {s}~{t}')
        for day, row in zip(d['dates'], d['rows']):
            v = list(row.values())[0]
            if v is not None:
                out[day] = v
        s = t + dt.timedelta(days=1)
    json.dump(out, open(fn, 'w'))
    return out


def deck(name):
    rows = list(csv.reader(open(f'{A}/{name}.csv')))
    hdr = rows[0][1:]
    data = {}
    for r in rows[1:]:
        data[r[0]] = [float(x) if x not in ('', 'None') else None for x in r[1:]]
    return hdr, data


def cum(series, base_day):
    b = series[base_day]
    return {k: v / b - 1 for k, v in series.items() if k >= base_day}


def mdd(series):
    ks = sorted(series)
    peak_k, peak = ks[0], series[ks[0]]
    best = (0, None, None)
    for k in ks:
        v = series[k]
        if v > peak:
            peak, peak_k = v, k
        dd = v / peak - 1
        if dd < best[0]:
            best = (dd, peak_k, k)
    return best


def compare(label, deck_series, fresh_cum, tol=0.0001):
    """deck_series: {date: value}（图里是四位小数的累计涨跌幅）"""
    n = bad = 0
    worst = (0, None)
    miss = 0
    for k, v in deck_series.items():
        if v is None:
            continue
        if k not in fresh_cum:
            miss += 1
            continue
        n += 1
        diff = abs(v - fresh_cum[k])
        if diff > worst[0]:
            worst = (diff, k)
        if diff > tol + 1e-9:
            bad += 1
    print(f'  {label}: 对照 {n} 点，超差 {bad}，最大差 {worst[0]:.5f}@{worst[1]}，图有而源无 {miss}')
    return bad == 0 and miss == 0


R = {}

# ---- P2 / P7 ----
print('P2/P7 2025-12-31~2026-11-13 全收益')
codes = {'中证A500全收益': '000510CNY010.CSI', '科创50全收益': '000688CNY01.SH',
         '中证红利低波动全收益': 'H20269.CSI', '创业板指全收益': '399606.SZ'}
hdr, p2 = deck('deck_Slide_g0')
hdr7, p7 = deck('deck_Slide-5410497_g0')
for name, code in codes.items():
    s = fetch(code, '2025-12-31', '2026-11-13')
    c = cum(s, '2025-12-31')
    i = hdr.index(name)
    compare('P2 ' + name, {k: v[i] for k, v in p2.items()}, c)
    if name in hdr7:
        j = hdr7.index(name)
        compare('P7 ' + name, {k: v[j] for k, v in p7.items()}, c)
    m = mdd(s)
    R[name] = (c['2026-11-13'], m)
    print(f'    累计 {c["2026-11-13"]:.2%}  最大回撤 {m[0]:.2%} {m[1]}→{m[2]}')

# ---- P11 / P12 ----
for slide, a, b in [('deck_Slide-5420639_g0', '2024-09-02', '2026-06-30'),
                    ('deck_Slide-5421867_g0', '2026-06-30', '2026-11-13')]:
    print(f'{slide} {a}~{b}')
    h, dd = deck(slide)
    for name in h:
        s = fetch(codes[name], '2024-09-02', '2026-11-13')
        c = cum({k: v for k, v in s.items() if k <= b}, a)
        compare(name, {k: v[h.index(name)] for k, v in dd.items()}, c)
        print(f'    {name} 区间 {c[b]:.2%}')

# ---- P13 中证全指 + 两融 ----
print('P13 中证全指（点位）/ 两融余额')
h, dd = deck('deck_Slide-5416467_g1')
s = fetch('000985.CSI', '2024-09-02', '2026-11-13')
compare('中证全指点位', {k: v[0] for k, v in dd.items()}, s, tol=0.006)
h, mm = deck('deck_Slide-5416467_g0')
ms = {}
for f in ['data/margin_wind.json', 'data/margin_wind_1116.json']:
    w = json.load(open(f))
    if 'content' in w:
        w = json.loads(w['content'][0]['text'])
    met = w['metrics'][0]
    for k, v in zip(met['date'], met['value']):
        k = k.replace('-', '')[:8]
        ms[f'{k[:4]}-{k[4:6]}-{k[6:]}'] = v
compare('两融余额(亿元)', {k: v[0] for k, v in mm.items()}, ms, tol=0.01)
mk = max(mm, key=lambda k: mm[k][0] or 0)
print(f'    两融高点 {mk} {mm[mk][0]}，末值 {max(mm)} {mm[max(mm)][0]}')

# ---- P14 300价值/成长 ----
print('P14 2004-12-31~2026-11-13 价格指数')
h, dd = deck('deck_Slide-5470098_g0')
for name, code in {'沪深300价值': '000919.CSI', '沪深300成长': '000918.CSI'}.items():
    s = fetch(code, '2004-12-31', '2026-11-13')
    c = cum(s, '2004-12-31')
    compare(name, {k: v[h.index(name)] for k, v in dd.items()}, c)
    yrs = (dt.date(2026, 11, 13) - dt.date(2004, 12, 31)).days / 365.25
    m = mdd(s)
    print(f'    累计 {c["2026-11-13"]:.2%} 年化 {(1 + c["2026-11-13"]) ** (1 / yrs) - 1:.2%} 回撤 {m[0]:.2%} {m[1]}→{m[2]}')

# ---- P18 主动优选 ----
print('P18 2026-06-18~2026-11-06')
h, dd = deck('deck_Slide-5435850_g0')
for name, code in {'沪深300': '000300.SH', '中证800': '000906.CSI', '中证全指': '000985.CSI'}.items():
    s = fetch(code, '2026-06-18', '2026-11-13')
    c = cum({k: v for k, v in s.items() if k <= '2026-11-06'}, '2026-06-18')
    compare(name, {k: v[h.index(name)] for k, v in dd.items()}, c)
    print(f'    {name} {c["2026-11-06"]:.2%}')
fn = f'{A}/src_nav.json'
if not os.path.exists(fn):
    r = subprocess.run(['python3', D, 'get', '/portfolio-nav', 'portfolioName=螺丝钉金钉宝主动优选C',
                        'beginDate=2026-06-18', 'endDate=2026-11-16', '--out', fn], capture_output=True, text=True)
nav = json.load(open(fn))
items = nav['data']['items'] if 'data' in nav else nav['items']
ns = {it['tradeDate']: it['unitNav'] for it in items}
c = cum(ns, '2026-06-18')
compare('主动优选', {k: v[h.index('主动优选')] for k, v in dd.items()}, c)
print(f'    主动优选 {c["2026-11-06"]:.2%}，工具箱最新净值日 {max(ns)}')

# ---- P10 中证A500 价格（定投打点图） ----
print('P10 中证A500 价格 2025-02-11~2026-11-13')
h, dd = deck('deck_Slide-5418988_g1')
s = fetch('000510.CSI', '2025-02-11', '2026-11-13')
c = cum(s, '2025-02-11')
compare('中证A500', {k: v[h.index('中证A500')] for k, v in dd.items()}, c, tol=0.0006)

# ---- P21 回测 ----
print('P21 2004-12-31~2025-12-31 回测（每年末再平衡 50/50）')
a5 = fetch('000510CNY010.CSI', '2004-12-31', '2025-12-31')
hl = fetch('H00922.CSI', '2004-12-31', '2025-12-31')
hs = fetch('H00300.CSI', '2004-12-31', '2025-12-31')
days = sorted(set(a5) & set(hl))
port = {}
va = vb = 0.5
la, lb = a5[days[0]], hl[days[0]]
for i, k in enumerate(days):
    va *= a5[k] / la
    vb *= hl[k] / lb
    la, lb = a5[k], hl[k]
    port[k] = va + vb
    if i + 1 < len(days) and days[i + 1][:4] != k[:4]:
        t = va + vb
        va = vb = t / 2
pc = cum(port, days[0])
hc = cum(hs, '2004-12-31')
h, dd = deck('deck_Slide-5413337_g0')
compare('组合', {k: v[0] for k, v in dd.items()}, pc, tol=0.01)
compare('沪深300全收益', {k: v[1] for k, v in dd.items()}, hc, tol=0.01)
print(f'    组合 收益 {pc["2025-12-31"]:.2%} 回撤 {mdd(port)[0]:.2%} {mdd(port)[1:]}')
print(f'    沪深300全收益 收益 {hc["2025-12-31"]:.2%} 回撤 {mdd(hs)[0]:.2%} {mdd(hs)[1:]}')
