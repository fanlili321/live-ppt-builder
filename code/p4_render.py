# -*- coding: utf-8 -*-
"""P4 美元降息页：双Y轴(美元指数线 + 国债收益率面积) + 加息/降息打点 + 箭头。matplotlib渲染PNG。"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import matplotlib.patches as mpatches
import matplotlib.font_manager as fm
from datetime import datetime

FONT = 'Source Han Sans CN'
matplotlib.rcParams['font.family'] = FONT
matplotlib.rcParams['axes.unicode_minus'] = False
ORANGE = '#F17640'; AREA = '#9DBFE6'; BORDER = '#FFBF00'
GREEN = '#2BA84A'; RED = '#FF2700'; DARK = '#333333'; GRAY = '#999999'


def render_p4(out_path, usd, yld, marks=None):
    """usd/yld: [{date,close}]; marks: [{date,label,type:'hike'/'cut'}]（可空）"""
    du = [datetime.strptime(x['date'], '%Y-%m-%d') for x in usd]
    vu = [x['close'] for x in usd]
    dy = [datetime.strptime(x['date'], '%Y-%m-%d') for x in yld]
    vy = [x['close'] for x in yld]

    fig, ax = plt.subplots(figsize=(20, 8.2), dpi=150)
    fig.patch.set_facecolor('white'); ax.set_facecolor('white')
    ax2 = ax.twinx()

    # 右轴：国债收益率 面积
    ax2.fill_between(dy, 0, vy, color=AREA, alpha=0.55, zorder=1)
    ax2.set_ylim(0, max(vy) * 1.15)
    # 左轴：美元指数 线
    ax.plot(du, vu, color=ORANGE, linewidth=4, zorder=4)
    ax.set_ylim(min(vu) * 0.96, max(vu) * 1.04)

    # 打点（在收益率面积上）
    ymap = {x['date']: x['close'] for x in yld}
    for m in (marks or []):
        if m['date'] not in ymap:
            continue
        d = datetime.strptime(m['date'], '%Y-%m-%d'); y = ymap[m['date']]
        col = RED if m.get('type') == 'hike' else GREEN
        ax2.scatter([d], [y], s=240, facecolor='white', edgecolor=col,
                    linewidth=3, zorder=6)
        ax2.annotate(m['label'], (d, y), xytext=(0, 12), textcoords='offset points',
                     ha='center', fontsize=20, fontweight='bold', color=DARK, fontfamily=FONT)

    # 样式
    for sp in ax.spines.values(): sp.set_visible(False)
    for sp in ax2.spines.values(): sp.set_visible(False)
    ax.yaxis.grid(True, color='#E0E0E0', linewidth=0.8, zorder=0); ax.set_axisbelow(True)
    ax.tick_params(axis='both', length=0, labelsize=26)
    ax2.tick_params(axis='y', length=0, labelsize=26)
    ax2.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{x:.0f}%'))
    for t in ax.get_yticklabels() + ax2.get_yticklabels(): t.set_fontfamily(FONT)
    # X 轴只首尾
    ax.set_xlim(du[0], du[-1])
    ax.set_xticks([du[0], du[-1]])
    ax.set_xticklabels([usd[0]['date'], usd[-1]['date']], fontsize=26, fontweight='bold', fontfamily=FONT)
    ax.get_xticklabels()[0].set_ha('left'); ax.get_xticklabels()[1].set_ha('right')

    # 图例（左上）：美元指数 / 降息 / 加息 / 国债收益率
    h = [plt.Line2D([0],[0], color=ORANGE, lw=4, label='美元指数（衡量美元汇率变化）'),
         plt.Line2D([0],[0], marker='o', mfc='white', mec=GREEN, mew=3, ls='', ms=13, label='降息'),
         plt.Line2D([0],[0], marker='o', mfc='white', mec=RED, mew=3, ls='', ms=13, label='加息'),
         mpatches.Patch(facecolor=AREA, alpha=0.55, label='美股10年期国债收益率')]
    ax.legend(handles=h, loc='upper left', bbox_to_anchor=(0, 1.12), frameon=False,
              ncol=4, fontsize=22, columnspacing=1.5, prop={'family': FONT, 'size': 22})

    fig.tight_layout(rect=[0.01, 0.04, 0.99, 0.99])
    fig.canvas.draw(); r = fig.canvas.get_renderer()
    from matplotlib.transforms import Bbox
    bb = ax.get_tightbbox(r); inv = fig.transFigure.inverted(); bf = bb.transformed(inv)
    pad = 0.02
    fig.add_artist(mpatches.FancyBboxPatch((bf.x0-pad, bf.y0-pad), bf.width+2*pad, bf.height+2*pad,
                   transform=fig.transFigure, boxstyle='square,pad=0', linewidth=4,
                   edgecolor=BORDER, facecolor='none', zorder=10))
    fig.text(bf.x0+bf.width+pad, bf.y0-pad-0.01, '数据来源：Wind', ha='right', va='top',
             color=GRAY, fontproperties=fm.FontProperties(family=FONT, size=22))
    fig.savefig(out_path, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    return out_path
