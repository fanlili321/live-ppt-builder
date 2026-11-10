#!/usr/bin/env python3
"""
银行螺丝钉直播PPT - 走势图生成器
样式完全对照大纲截图参考图（img_03/04/15/16）
"""

import json, subprocess, os
from datetime import datetime
from io import BytesIO

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import matplotlib.font_manager as fm

# 设置中文字体
plt.rcParams['font.family'] = ['Heiti TC', 'Hiragino Sans GB', 'Arial Unicode MS', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

# ── 颜色 ────────────────────────────────────────────────
LINE_ORANGE  = '#E97C2D'   # 中证全指主线
LINE_GRAY    = '#929292'   # 全收益副线
GRID_GRAY    = '#E8E8E8'   # 横向网格线
DOT_RED      = '#CC0000'   # 标注红圈
BORDER_YELLOW= '#FFC000'   # 图表外框（在 PPT 里叠加，不在 matplotlib 里画）
TEXT_DARK    = '#333333'
TEXT_GRAY    = '#999999'

# ── Wind 数据获取 ────────────────────────────────────────
WIND_DIR = os.path.expanduser('~/.agents/skills/wind-mcp-skill')

def _to_weekly(rows: list[dict]) -> list[dict]:
    """日频→周频：每个ISO周保留最后一个交易日的收盘。"""
    import datetime as _dt
    wk = {}
    for r in rows:
        y, w, _ = _dt.date.fromisoformat(r['date']).isocalendar()
        wk[(y, w)] = r            # 同周后出现的覆盖前面=本周最后一天
    return [wk[k] for k in sorted(wk)]


def _fetch_raw(windcode, start, end):
    """单段抓取（含重试）。"""
    import time
    params = json.dumps({"windcode": windcode,
                         "begin_date": start.replace('-',''),
                         "end_date":   end.replace('-','')})
    for attempt in range(6):
        r = subprocess.run(
            ['node', 'scripts/cli.mjs', 'call', 'index_data', 'get_index_kline', params],
            capture_output=True, text=True, cwd=WIND_DIR, timeout=90
        )
        stdout = r.stdout.strip()
        if not stdout:
            time.sleep(3)
            continue
        try:
            outer = json.loads(stdout)
            inner = json.loads(outer['content'][0]['text'])
            rows  = inner['data']['rows']
            return [{'date': row[0][:10], 'close': float(row[2])} for row in rows]
        except Exception as e:
            print(f'  [Wind重试 {attempt+1}] {e}')
            time.sleep(3)
    raise RuntimeError(f'Wind数据获取失败: {windcode} {start}~{end}')


def fetch_kline(windcode: str, start: str, end: str, freq: str = 'D') -> list[dict]:
    """获取K线 [{'date','close'},...]。freq='D'日频/'W'周频。长区间按年分段抓取再合并（更稳）。"""
    import datetime as _dt
    s = _dt.date.fromisoformat(start); e = _dt.date.fromisoformat(end)
    data, seen = [], set()
    cur = s
    while cur <= e:
        seg_end = min(_dt.date(cur.year, 12, 31), e)
        for row in _fetch_raw(windcode, cur.isoformat(), seg_end.isoformat()):
            if row['date'] not in seen:
                seen.add(row['date']); data.append(row)
        cur = _dt.date(cur.year + 1, 1, 1)
    data.sort(key=lambda x: x['date'])
    return _to_weekly(data) if str(freq).upper().startswith('W') else data


def find_hi_lo(data: list[dict]) -> tuple[dict, dict]:
    """在列表中找最高点和最低点"""
    hi = max(data, key=lambda x: x['close'])
    lo = min(data, key=lambda x: x['close'])
    return hi, lo


# ── 核心图表绘制 ─────────────────────────────────────────
def draw_trend_chart(
    data:              list[dict],           # 主线数据
    highlight:         list[dict] = None,    # 红圈标注点 [{'date','close','offset':(dx,dy)}]
    color_bands:       list[dict] = None,    # 彩色背景 [{'start','end','color','alpha'}]
    second_data:       list[dict] = None,    # 第二条线
    second_label:      str = '中证全指全收益',
    first_label:       str = '中证全指',
    bubble_labels:     list[dict] = None,    # 气泡 [{'date','close','text','offset':(dx,dy)}]
    trend_arrow:       bool = False,         # 是否画趋势箭头
    y_pct:             bool = False,         # Y轴百分比格式
    figsize:           tuple = (11.0, 4.8),
) -> bytes:
    """生成图表 PNG bytes（不含黄色外框，外框在 PPT 里加）"""

    dates  = [datetime.strptime(d['date'], '%Y-%m-%d') for d in data]
    closes = [d['close'] for d in data]

    fig, ax = plt.subplots(figsize=figsize, facecolor='white')
    fig.subplots_adjust(left=0.07, right=0.97, top=0.93, bottom=0.10)

    # ── 彩色区间 ──
    if color_bands:
        for b in color_bands:
            d1 = datetime.strptime(b['start'], '%Y-%m-%d')
            d2 = datetime.strptime(b['end'],   '%Y-%m-%d')
            ax.axvspan(d1, d2, alpha=b.get('alpha', 0.15), color=b['color'], zorder=1)

    # ── 主线 ──
    ax.plot(dates, closes, color=LINE_ORANGE, linewidth=2.2, zorder=3,
            label=first_label)

    # ── 第二条线 ──
    if second_data:
        d2 = [datetime.strptime(d['date'], '%Y-%m-%d') for d in second_data]
        c2 = [d['close'] for d in second_data]
        ax.plot(d2, c2, color=LINE_GRAY, linewidth=1.8, zorder=3, label=second_label)

    # ── 图例 ──
    if second_data or (first_label and first_label != '中证全指'):
        leg = ax.legend(loc='upper left', fontsize=11, frameon=False,
                        handlelength=1.8, handletextpad=0.6,
                        labelcolor=TEXT_DARK)

    # ── 样式 ──
    ax.set_facecolor('white')
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.yaxis.grid(True,  color=GRID_GRAY, linewidth=0.6, zorder=0)
    ax.xaxis.grid(False)
    ax.tick_params(axis='both', which='both', length=0)

    # ── X轴：只显示起止日期 ──
    ax.set_xlim(dates[0], dates[-1])
    start_str = dates[0].strftime('%Y-%m-%d')
    end_str   = dates[-1].strftime('%Y-%m-%d')
    ax.set_xticks([dates[0], dates[-1]])
    ax.set_xticklabels([start_str, end_str],
                        fontsize=11, color=TEXT_DARK, fontweight='bold')
    ax.get_xticklabels()[0].set_ha('left')
    ax.get_xticklabels()[1].set_ha('right')

    # ── Y轴格式 ──
    if y_pct:
        def pct_fmt(x, _):
            return f'{x:.0f}%' if x == int(x) else f'{x:.1f}%'
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(pct_fmt))
    else:
        def num_fmt(x, _):
            if x >= 1000: return f'{int(x):,}'
            return f'{x:.0f}' if x == int(x) else f'{x:.1f}'
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(num_fmt))
    ax.tick_params(axis='y', labelsize=10, labelcolor='#555555')

    # ── 红圈标注 ──
    if highlight:
        y_min, y_max = min(closes), max(closes)
        y_range = y_max - y_min or 1
        for pt in highlight:
            pd_ = datetime.strptime(pt['date'], '%Y-%m-%d')
            pc_ = pt['close']
            ax.plot(pd_, pc_, 'o', color=DOT_RED, markersize=9,
                    markerfacecolor=DOT_RED, zorder=6)
            # 自动判断文字方向：避免超出图框
            if 'offset' in pt:
                dx, dy = pt['offset']
            else:
                # 垂直方向：下半部分往上，上半部分往下
                y_rel = (pc_ - y_min) / y_range  # 0=底 1=顶
                dy = 14 if y_rel < 0.5 else -38
                # 水平方向：右1/3往左，否则往右
                x_rel = (pd_ - dates[0]).days / max((dates[-1] - dates[0]).days, 1)
                dx = -90 if x_rel > 0.65 else 12
            ha = 'left' if dx >= 0 else 'right'
            va = 'bottom' if dy >= 0 else 'top'
            label = f"{pt['date']}\n收盘价{pc_:.2f}"
            ax.annotate(label, xy=(pd_, pc_),
                        xytext=(dx, dy), textcoords='offset points',
                        fontsize=10, fontweight='bold', color=TEXT_DARK,
                        ha=ha, va=va, zorder=7, clip_on=False)

    # ── 气泡标注 ──
    if bubble_labels:
        for bub in bubble_labels:
            bd = datetime.strptime(bub['date'], '%Y-%m-%d')
            bc = bub['close']
            dx, dy = bub.get('offset', (0, -40))
            ax.annotate(bub['text'], xy=(bd, bc),
                        xytext=(dx, dy), textcoords='offset points',
                        fontsize=9, ha='center', va='top', color=TEXT_DARK,
                        bbox=dict(boxstyle='round,pad=0.4',
                                  facecolor='#FFDDC4',
                                  edgecolor='#E97C2D', linewidth=1.0),
                        zorder=7)

    # ── 趋势箭头（市场底部上涨页） ──
    if trend_arrow and len(dates) >= 2:
        y_lo = min(closes); y_hi = max(closes)
        y_mid = y_lo + (y_hi - y_lo) * 0.25
        ax.annotate('', xy=(dates[-1], y_hi * 0.85),
                    xytext=(dates[0], y_mid),
                    arrowprops=dict(arrowstyle='->', color=TEXT_DARK,
                                   lw=2.0, mutation_scale=20))

    buf = BytesIO()
    plt.savefig(buf, format='png', dpi=150,
                facecolor='white', edgecolor='none')
    plt.close(fig)
    buf.seek(0)
    return buf.read()


# ── 快速测试 ─────────────────────────────────────────────
if __name__ == '__main__':
    print('拉取第一轮牛市数据：中证全指 2004-12-31 ~ 2007-10-31')
    data = fetch_kline('000985.CSI', '2004-12-31', '2007-10-31')
    print(f'  共 {len(data)} 条，首: {data[0]}，尾: {data[-1]}')

    hi, lo = find_hi_lo(data)
    print(f'  最高点: {hi}')
    print(f'  最低点: {lo}')

    img = draw_trend_chart(
        data,
        highlight=[
            {'date': lo['date'], 'close': lo['close']},
            {'date': hi['date'], 'close': hi['close']},
        ],
    )
    out = '/tmp/sample_chart.png'
    with open(out, 'wb') as f:
        f.write(img)
    print(f'  样图已保存: {out}')
