#!/usr/bin/env python3
"""
走势图作图引擎 —— 复刻 PPT-chart 工具(chart_app.py build_figure)的格式：
matplotlib；橙线 #F17640 线宽4；白底；灰网格 #E0E0E0；图例左上无边框；
X轴只显首尾日期加粗；黄色外框 #FFBF00 紧包；来源右下灰字；思源黑体；
高低点打红点并标注 日期+收盘价。输出 PNG 路径。
"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import matplotlib.patches as mpatches
import matplotlib.font_manager as fm
import numpy as np
import pandas as pd
from matplotlib.transforms import Bbox

FONT   = 'Source Han Sans CN'
FS     = 30
ORANGE = '#F17640'
GRAY   = '#929292'
BORDER = '#FFBF00'
RED    = '#FF2700'
PAD_H, PAD_V, SOURCE_GAP = 0.022, 0.022, 0.010
matplotlib.rcParams['font.family'] = FONT
matplotlib.rcParams['axes.unicode_minus'] = False


def _nice_ticks(vals, n=5):
    lo, hi = float(np.min(vals)), float(np.max(vals))
    if hi == lo:
        hi += 1
    loc = mticker.MaxNLocator(nbins=n, prune=None)
    ticks = loc.tick_values(lo, hi)
    pad = (hi - lo) * 0.08
    return ticks, lo - pad, hi + pad


def render_chart(out_path, series, source='数据来源：Wind',
                 mode='points', hilo_series=0, marks=None):
    """
    series: [{'name','color','dates':[...],'values':[...]}]
    mode  : 'points'（指数点数）| 'percent'（已是百分比，直接显示带%）
    hilo_series: 在第几条线上标红点高低（None 不标）
    marks : 自定义点 [{'date','value','label'}] 覆盖默认高低点
    """
    fig, ax = plt.subplots(figsize=(20, 9), dpi=150)
    fig.patch.set_facecolor('white'); ax.set_facecolor('white')

    all_dates = None
    pseries = []
    for s in series:
        idx = pd.to_datetime(s['dates'])
        ser = pd.Series(s['values'], index=idx).sort_index()
        pseries.append({**s, 'ser': ser})
        all_dates = ser.index if all_dates is None else all_dates.union(ser.index)

    for s in pseries:
        ax.plot(s['ser'].index, s['ser'].values, color=s['color'],
                linewidth=4, zorder=3, solid_capstyle='round')

    # X 轴范围 + 仅首尾日期
    rng = (all_dates[-1] - all_dates[0]).days
    pad = pd.Timedelta(days=max(int(rng * 0.015), 3))
    ax.set_xlim(all_dates[0] - pad, all_dates[-1] + pad)
    ax.set_xticks([all_dates[0], all_dates[-1]])
    ax.set_xticklabels([all_dates[0].strftime('%Y-%m-%d'),
                        all_dates[-1].strftime('%Y-%m-%d')],
                       fontsize=FS, fontweight='bold', fontfamily=FONT)
    ax.tick_params(axis='x', length=0, pad=10)

    # Y 轴
    allv = np.concatenate([s['ser'].values for s in pseries])
    ticks, ylo, yhi = _nice_ticks(allv)
    ax.set_ylim(ylo, yhi)
    ax.yaxis.set_major_locator(mticker.FixedLocator(ticks))
    if mode == 'percent':
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(
            lambda x, _: f'{int(round(x))}%'))
    else:
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(
            lambda x, _: f'{int(round(x)):,}'))
    ax.tick_params(axis='y', labelsize=FS, length=0, pad=10, left=False)
    for t in ax.get_yticklabels():
        t.set_fontfamily(FONT); t.set_fontsize(FS)

    ax.yaxis.grid(True, color='#E0E0E0', linewidth=0.8, zorder=1)
    ax.xaxis.grid(False); ax.set_axisbelow(True)
    for sp in ax.spines.values():
        sp.set_visible(False)

    # 红点高低标注
    if hilo_series is not None and hilo_series < len(pseries):
        ser = pseries[hilo_series]['ser']
        pts = marks if marks else [
            {'date': ser.idxmin().strftime('%Y-%m-%d'), 'value': float(ser.min())},
            {'date': ser.idxmax().strftime('%Y-%m-%d'), 'value': float(ser.max())},
        ]
        for p in pts:
            d = pd.to_datetime(p['date']); v = p['value']
            ax.scatter([d], [v], s=260, facecolor=RED, edgecolor='white',
                       linewidth=2, zorder=6)
            lbl = p.get('label', f"{p['date']}\n收盘价{v:,.2f}")
            up = v < (ylo + yhi) / 2
            ax.annotate(lbl, (d, v),
                        xytext=(0, 28 if up else -28),
                        textcoords='offset points',
                        ha='center', va='bottom' if up else 'top',
                        fontsize=FS - 4, fontweight='bold', color='#333333',
                        fontfamily=FONT)

    # 图例（左上、无框）
    handles = [plt.Line2D([0], [0], color=s['color'], linewidth=4, label=s['name'])
               for s in pseries]
    fig.tight_layout(rect=[0.0, 0.10, 1.0, 1.0])
    leg = ax.legend(handles=handles, loc='upper left',
                    bbox_to_anchor=(0, 1.10), frameon=False,
                    ncol=len(handles), handlelength=1.8, columnspacing=1.2,
                    prop={'family': FONT, 'size': FS})

    fig.canvas.draw()
    r = fig.canvas.get_renderer()
    full = Bbox.union([ax.get_tightbbox(r), leg.get_window_extent(r)])
    inv = fig.transFigure.inverted(); bf = full.transformed(inv)
    bx, by = bf.x0 - PAD_H, bf.y0 - PAD_V
    bw, bh = bf.width + 2 * PAD_H, bf.height + 2 * PAD_V
    fig.add_artist(mpatches.FancyBboxPatch(
        (bx, by), bw, bh, transform=fig.transFigure,
        boxstyle='square,pad=0', linewidth=4,
        edgecolor=BORDER, facecolor='none', zorder=10))
    fig.text(bx + bw, by - SOURCE_GAP, source, ha='right', va='top',
             color='#999999',
             fontproperties=fm.FontProperties(family=FONT, size=FS))

    fig.savefig(out_path, format='png', dpi=150,
                bbox_inches='tight', facecolor='white')
    plt.close(fig)
    return out_path
