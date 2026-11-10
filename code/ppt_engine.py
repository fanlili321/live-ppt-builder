#!/usr/bin/env python3
"""
453期直播PPT生成器 v5
规则：
- 字号按母版原生（PPT尺寸）: 封面60pt / 页标题44pt / 正文24pt，自动适配
- 标题颜色 = 主色调 #F17640（不是黑色）
- 正文文字黑色左对齐
- 图例靠左，X轴日期在黄框内，无倒影/阴影效果
- 文字严格按大纲，不自行添加样式
"""

import os
import sys

CODE_DIR = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(CODE_DIR)
sys.path.insert(0, CODE_DIR)
from chart_gen import fetch_kline, find_hi_lo
from datetime import datetime

from pptx import Presentation
from pptx.util import Pt, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.chart import (XL_CHART_TYPE, XL_LEGEND_POSITION,
                             XL_MARKER_STYLE, XL_LABEL_POSITION)
from pptx.chart.data import ChartData
from pptx.oxml.ns import qn
from pptx.enum.shapes import MSO_SHAPE
from lxml import etree
from PIL import Image

IMG_DIR = os.path.join(SKILL_DIR, 'examples', 'imgs_456')

# ── 颜色 ─────────────────────────────────────────────────
C_YELLOW = RGBColor(0xFF, 0xBF, 0x00)   # 走势图黄框 #FFBF00（保持不变）
C_BLACK  = RGBColor(0x00, 0x00, 0x00)
C_DARK   = RGBColor(0x33, 0x33, 0x33)
C_GRAY   = RGBColor(0x99, 0x99, 0x99)
C_ORANGE = RGBColor(0xF1, 0x76, 0x40)
C_MAIN   = RGBColor(0xF1, 0x76, 0x40)   # 品牌主色 #F17640（与模版主题色统一），标题/箭头要点用色
C_LINE   = RGBColor(0xF1, 0x76, 0x40)   # 走势图橙线 #F17640（PPT-chart）
C_RED    = RGBColor(0xFF, 0x27, 0x00)   # 高低点红圆点
# 多曲线配色顺序（照搬 tool-ppt-chart 的 AUTO_COLORS）
PALETTE  = ['F17640', '7CA7E0', '929292', '6674A6', 'FF2700', 'C24785', 'C8A028']
def _pal(i):
    h = PALETTE[i % len(PALETTE)]
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))
# 样式（按图表缩放比 ≈0.63 从 tool 的 matplotlib 值映射）
CHART_LINE_PT = 2.5    # tool linewidth 4 → ~2.5pt
CHART_FRAME_PT = 2.5   # tool 黄框 linewidth 4
CHART_FONT_PT = 18     # tool FONT_SIZE 30 → ~18pt

FONT = 'Source Han Sans CN Regular'

WINDCODE  = '000985.CSI'
TMPL_PATH = '/template/母版（ppt格式）.pptx'
TMPL_PATH = os.path.join(SKILL_DIR, 'template', '母版（ppt格式）.pptx')
OUT_PATH  = os.path.expanduser('~/Downloads/453期直播PPT.pptx')

# ── 默认字号（可自动缩小，但不偏离默认值太多）────────────
COVER_PT  = 60   # 封面标题 60pt（规范：封面60-80pt橙色）
TITLE_PT  = 40   # 内容页主标题 40pt（规范：内容页36-48pt）
BODY_PT   = 20   # 正文/要点 20pt（用户批注：所有正文文字缩小为20）
MIN_RATIO = 0.65  # 最大缩小到默认的65%

# ── 布局（Master 1，正文页）──────────────────────────────
# Master 1: 顶部淡黄条 y=0.8~2.3，logo(28.7,0.9)，版权 y=17.5
CONTENT_TOP    = 2.5    # 内容区开始 y（条带下方）
CONTENT_BOTTOM = 17.0   # 内容区结束 y（版权上方）
CONTENT_L      = 0.8    # 左边距
CONTENT_W      = 32.0   # 内容宽度（总宽 33.87）

# 图表区域（图表左右边距与内容一致，高度动态计算）
CHART_L = 0.8
CHART_W = 32.0

# ── 工具：字号自动适配 ───────────────────────────────────
def auto_pt(text: str, default_pt: int) -> int:
    """
    根据文字数量自动缩小字号，不低于 default_pt * MIN_RATIO。
    阶梯式缩减，避免偏离默认值太大。
    """
    n = len(text.replace('\n', '').replace(' ', ''))
    if   n <= 12: ratio = 1.00
    elif n <= 20: ratio = 0.88
    elif n <= 30: ratio = 0.78
    elif n <= 45: ratio = 0.70
    else:         ratio = MIN_RATIO
    return max(round(default_pt * ratio), round(default_pt * MIN_RATIO))

def title_fit(text: str) -> int:
    """正文标题固定 40pt（规范E2）；文字过多时才适当缩小。"""
    n = len(text.replace('\n', '').replace(' ', ''))
    if n <= 24: return TITLE_PT      # 默认 40
    if n <= 30: return 36
    return 32

def line_height_cm(pt: int) -> float:
    """pt → 行高 cm（含 1.2 倍行距）"""
    return pt * (2.54 / 72) * 1.2

def estimate_lines(text: str, pt: int, box_w_cm: float) -> int:
    """估算文字在指定宽度内需要几行"""
    char_w = pt * (2.54 / 72)
    chars_per_line = max(1, box_w_cm / char_w)
    manual_lines = text.count('\n') + 1
    total_chars = len(text.replace('\n', ''))
    wrap_lines = -(-total_chars // max(1, int(chars_per_line)))  # ceiling division
    return max(manual_lines, wrap_lines)

def text_height_cm(text: str, pt: int, box_w_cm: float) -> float:
    """估算文字块高度 cm"""
    n = estimate_lines(text, pt, box_w_cm)
    return line_height_cm(pt) * n + 0.3  # +0.3 padding

# ── 模版工具 ──────────────────────────────────────────────
def get_layout(prs, master_idx, layout_name=None):
    m = list(prs.slide_masters)[master_idx]
    if layout_name:
        for lyt in m.slide_layouts:
            if lyt.name == layout_name:
                return lyt
    return m.slide_layouts[0]

def remove_slide(prs, idx):
    sldIdLst = prs.slides._sldIdLst
    sld = sldIdLst[idx]
    rId = sld.get(
        '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
    prs.part.drop_rel(rId)
    sldIdLst.remove(sld)

# ── 文字工具 ──────────────────────────────────────────────
def set_run_font(run, name=FONT):
    """同时设置西文(latin)、中文(ea)、复杂文种(cs)字体，确保中文用思源黑体"""
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    for tag in ('a:latin', 'a:ea', 'a:cs'):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set('typeface', name)

def add_tb(slide, text, left, top, width, height,
           size=18, bold=False, color=C_BLACK, align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(Cm(left), Cm(top), Cm(width), Cm(height))
    tf = tb.text_frame
    tf.word_wrap = True
    # 关键：显式禁用自动收缩。不写这行时 pptx 转 Keynote 后文本框会被
    # Keynote 按"缩小边框贴合文字"处理，导致框比设定的 width 窄很多、
    # 内容偏左、右侧大片空白（PDF 导出走标称尺寸看不出来，实机渲染才会露出）。
    from pptx.enum.text import MSO_AUTO_SIZE
    tf.auto_size = MSO_AUTO_SIZE.NONE
    for i, line in enumerate(text.split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = line
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
        set_run_font(run)
    return tb

def set_ph_text(sl, text, size, bold=False, color=C_BLACK, align=PP_ALIGN.LEFT):
    """向 placeholder 写字（封面用）"""
    for ph in sl.placeholders:
        ph.text = text
        tf = ph.text_frame
        tf.word_wrap = True
        for para in tf.paragraphs:
            para.alignment = align
            for run in para.runs:
                run.font.size = Pt(size)
                run.font.bold = bold
                run.font.color.rgb = color
                set_run_font(run)
        return

# ── 图表 XML 工具 ─────────────────────────────────────────
C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

def _ensure(parent, tag):
    el = parent.find(tag)
    if el is None:
        el = etree.SubElement(parent, tag)
    return el

def _solid_fill(parent, hex_clr, ns):
    spPr = _ensure(parent, f'{{{ns}}}spPr')
    sf   = _ensure(spPr,   f'{{{A_NS}}}solidFill')
    sc   = _ensure(sf,     f'{{{A_NS}}}srgbClr')
    sc.set('val', hex_clr)

def _ln_solid(parent, hex_clr, ns):
    spPr = _ensure(parent, f'{{{ns}}}spPr')
    ln   = _ensure(spPr,   f'{{{A_NS}}}ln')
    sf   = _ensure(ln,     f'{{{A_NS}}}solidFill')
    sc   = _ensure(sf,     f'{{{A_NS}}}srgbClr')
    sc.set('val', hex_clr)

def _no_effects(parent, ns):
    """显式清空 effectLst（去掉倒影/阴影）"""
    spPr = _ensure(parent, f'{{{ns}}}spPr')
    effectLst = spPr.find(f'{{{A_NS}}}effectLst')
    if effectLst is None:
        etree.SubElement(spPr, f'{{{A_NS}}}effectLst')
    else:
        for child in list(effectLst):
            effectLst.remove(child)

def style_chart_xml(chart, y_pct=False):
    """
    XML 精调：
    - 白色背景、无边框、无倒影
    - X 轴隐藏刻度标签（日期另行手动放入黄框内）
    - Y 轴数字格式
    - 灰色网格线
    - 无数据标记点
    """
    cs        = chart._chartSpace
    chartElem = cs.find(f'{{{C_NS}}}chart')
    plotArea  = chartElem.find(f'{{{C_NS}}}plotArea')

    # chartSpace 白色背景 + 无倒影
    _solid_fill(cs, 'FFFFFF', C_NS)
    _no_effects(cs, C_NS)

    # plotArea 白色背景 + 无边框
    _solid_fill(plotArea, 'FFFFFF', C_NS)
    spPr = _ensure(plotArea, f'{{{C_NS}}}spPr')
    ln   = _ensure(spPr, f'{{{A_NS}}}ln')
    if ln.find(f'{{{A_NS}}}noFill') is None:
        etree.SubElement(ln, f'{{{A_NS}}}noFill')

    def _ln_none(parent):
        spPr = _ensure(parent, f'{{{C_NS}}}spPr')
        ln   = _ensure(spPr, f'{{{A_NS}}}ln')
        for c in list(ln):
            ln.remove(c)
        etree.SubElement(ln, f'{{{A_NS}}}noFill')

    # X 轴：无刻度标签、无刻度线、无轴线（去掉多余竖线）
    for catAx in plotArea.findall(f'{{{C_NS}}}catAx'):
        for tag, val in [('tickLblPos','none'),
                         ('majorTickMark','none'),
                         ('minorTickMark','none')]:
            _ensure(catAx, f'{{{C_NS}}}{tag}').set('val', val)
        _ln_none(catAx)

    # Y 轴：数字格式、无刻度线、无轴线
    for valAx in plotArea.findall(f'{{{C_NS}}}valAx'):
        nf = _ensure(valAx, f'{{{C_NS}}}numFmt')
        nf.set('formatCode', '0"%"' if y_pct else '#,##0')
        nf.set('sourceLinked', '0')
        for tag in ('majorTickMark', 'minorTickMark'):
            _ensure(valAx, f'{{{C_NS}}}{tag}').set('val', 'none')
        _ln_none(valAx)

    # 主网格线浅灰
    for gl in plotArea.findall(f'.//{{{C_NS}}}majorGridlines'):
        _ln_solid(gl, 'E8E8E8', C_NS)
    # 注：标记点由各系列单独控制（主线 NONE、高低点 CIRCLE），此处不再统一清除


# ── 折线图（含黄框，X 轴日期在黄框内）───────────────────
def _legend_top_left(chart):
    """把图例固定到图表左上角（对齐 tool-ppt-chart 的 upper-left）。"""
    legend = chart._chartSpace.find(f'.//{{{C_NS}}}legend')
    if legend is None:
        return
    pos = legend.find(f'{{{C_NS}}}legendPos')
    layout = etree.SubElement(legend, f'{{{C_NS}}}layout')
    ml = etree.SubElement(layout, f'{{{C_NS}}}manualLayout')
    for tag, val in [('xMode', 'edge'), ('yMode', 'edge')]:
        e = etree.SubElement(ml, f'{{{C_NS}}}{tag}'); e.set('val', val)
    for tag, val in [('x', '0.02'), ('y', '0.02')]:
        e = etree.SubElement(ml, f'{{{C_NS}}}{tag}'); e.set('val', val)
    if pos is not None:           # layout 须紧跟 legendPos
        legend.remove(layout); pos.addnext(layout)


def _hide_legend_entry(chart, idx):
    """从图例中删除指定系列的条目（高低点标记系列不显示在图例里）。"""
    legend = chart._chartSpace.find(f'.//{{{C_NS}}}legend')
    if legend is None:
        return
    le = etree.SubElement(legend, f'{{{C_NS}}}legendEntry')
    # legendEntry 需在 legendPos 之后、overlay 之前——简单起见追加后排序由PPT容错
    i = etree.SubElement(le, f'{{{C_NS}}}idx'); i.set('val', str(idx))
    d = etree.SubElement(le, f'{{{C_NS}}}delete'); d.set('val', '1')
    # 移到 legendPos 之后
    legend.remove(le)
    pos = legend.find(f'{{{C_NS}}}legendPos')
    pos.addnext(le)


def _nice_bound(v, up):
    """把数值向上/向下取整到“漂亮”刻度，用于 Y 轴自适应边界。"""
    import math
    if v == 0:
        return 0
    mag = 10 ** (math.floor(math.log10(abs(v))) - 1)
    return (math.ceil(v / mag) if up else math.floor(v / mag)) * mag

def _nice_axis(vmin, vmax):
    """返回 (min, max, step) 漂亮整数刻度，范围包住 vmin/vmax 并留余量。"""
    import math
    rng = (vmax - vmin) or abs(vmax) or 1
    raw = rng / 4.0
    mag = 10 ** math.floor(math.log10(raw)) if raw > 0 else 1
    step = mag
    for m in (1, 2, 2.5, 5, 10):
        if raw <= m * mag:
            step = m * mag; break
    lo = math.floor(vmin / step) * step
    hi = math.ceil(vmax / step) * step
    if hi - vmax < step * 0.15:   # 顶部留点余量，最高点别贴边
        hi += step
    if vmin - lo < step * 0.15:
        lo -= step
    return lo, hi, step

# 绘图区在图表对象内的固定位置（占比）。PLOT_H 拉满，使曲线底部贴近X轴日期
PLOT_X, PLOT_Y, PLOT_W, PLOT_H = 0.11, 0.03, 0.86, 0.95
FRAME_MARGIN = 0.35   # 图例↔黄框顶 与 日期↔黄框底 的留白（上下一致）

def _set_plot_area(chart):
    pa = chart._chartSpace.find(f'.//{{{C_NS}}}plotArea')
    if pa is None:
        return
    for old in pa.findall(f'{{{C_NS}}}layout'):
        pa.remove(old)
    layout = pa.makeelement(f'{{{C_NS}}}layout', {})
    pa.insert(0, layout)
    ml = etree.SubElement(layout, f'{{{C_NS}}}manualLayout')
    t = etree.SubElement(ml, f'{{{C_NS}}}layoutTarget'); t.set('val', 'inner')
    for tag, v in [('xMode','edge'),('yMode','edge')]:
        e = etree.SubElement(ml, f'{{{C_NS}}}{tag}'); e.set('val', v)
    for tag, v in [('x',PLOT_X),('y',PLOT_Y),('w',PLOT_W),('h',PLOT_H)]:
        e = etree.SubElement(ml, f'{{{C_NS}}}{tag}'); e.set('val', str(v))

def _nudge_label(dl, dx, dy):
    """给数据标签加手动偏移(占图比例)，把文字从曲线上挪开。dy<0 上移。"""
    dLbl = dl._dLbl
    idx = dLbl.find(f'{{{C_NS}}}idx')
    layout = dLbl.makeelement(f'{{{C_NS}}}layout', {})
    if idx is not None:
        idx.addnext(layout)
    else:
        dLbl.insert(0, layout)
    ml = etree.SubElement(layout, f'{{{C_NS}}}manualLayout')
    for tag, v in [('x', dx), ('y', dy)]:
        e = etree.SubElement(ml, f'{{{C_NS}}}{tag}'); e.set('val', str(v))

def _set_chart_font(chart):
    """给整个图表设默认字体=思源黑体（含中文 ea），字号 CHART_FONT_PT。"""
    cs = chart._chartSpace
    # 移除已有 txPr 再重建
    for old in cs.findall(f'{{{C_NS}}}txPr'):
        cs.remove(old)
    txPr = etree.SubElement(cs, f'{{{C_NS}}}txPr')
    etree.SubElement(txPr, f'{{{A_NS}}}bodyPr')
    etree.SubElement(txPr, f'{{{A_NS}}}lstStyle')
    p = etree.SubElement(txPr, f'{{{A_NS}}}p')
    pPr = etree.SubElement(p, f'{{{A_NS}}}pPr')
    defRPr = etree.SubElement(pPr, f'{{{A_NS}}}defRPr')
    defRPr.set('sz', str(int(CHART_FONT_PT * 100)))
    for tag in ('latin', 'ea', 'cs'):
        e = etree.SubElement(defRPr, f'{{{A_NS}}}{tag}'); e.set('typeface', FONT)
    etree.SubElement(p, f'{{{A_NS}}}endParaRPr')


def _set_fill_alpha(shape, hex_clr, alpha_pct):
    """给形状设半透明纯色填充（alpha_pct: 0-100，越小越透明）。"""
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(int(hex_clr[0:2],16), int(hex_clr[2:4],16), int(hex_clr[4:6],16))
    sc = shape.fill._xPr.find(qn('a:solidFill')).find(qn('a:srgbClr'))
    for c in list(sc):
        sc.remove(c)
    a = sc.makeelement(qn('a:alpha'), {}); a.set('val', str(int(alpha_pct * 1000)))
    sc.append(a)


def add_line_chart(slide, data, chart_l, chart_t, chart_w, chart_h,
                   second_data=None, y_pct=False,
                   hi_pt=None, lo_pt=None, bands=None,
                   series_name='中证全指', second_name='中证全指全收益'):
    """原生折线图 + 黄色外框。图表对象在黄框内缩进留白；Y轴按数据自适应。"""
    # 双线时按共同交易日(交集)对齐，避免错位/断线
    if second_data:
        s2 = {d['date']: d['close'] for d in second_data}
        data = [d for d in data if d['date'] in s2]
    dates = [d['date'] for d in data]
    cd = ChartData()
    cd.categories = dates
    cd.add_series(series_name, [d['close'] for d in data])
    if second_data:
        cd.add_series(second_name, [s2[dt] for dt in dates])

    marks, marker_idx = [], None
    if hi_pt or lo_pt:
        vals = [None] * len(dates)
        for p in (lo_pt, hi_pt):
            if p and p['date'] in dates:
                i = dates.index(p['date']); vals[i] = p['close']; marks.append((i, p))
        if marks:
            cd.add_series('高低点', vals)
            marker_idx = len(cd._series) - 1

    # 黄框内：顶部图例带、底部日期带，上下留白一致(FRAME_MARGIN)
    SIDE = 0.7
    LEG_H, DATE_H = 0.70, 0.65
    TOP_BAND = FRAME_MARGIN + LEG_H            # 图例带
    BOT_BAND = FRAME_MARGIN + DATE_H + 0.15    # 日期带(含曲线↔日期小缝0.15)
    obj_l, obj_t = chart_l + SIDE, chart_t + TOP_BAND
    obj_w, obj_h = chart_w - 2*SIDE, chart_h - TOP_BAND - BOT_BAND
    cs_shape = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE_MARKERS if marks else XL_CHART_TYPE.LINE,
        Cm(obj_l), Cm(obj_t), Cm(obj_w), Cm(obj_h), cd)
    ch = cs_shape.chart

    ch.series[0].format.line.color.rgb = _pal(0)
    ch.series[0].format.line.width = Pt(CHART_LINE_PT)
    ch.series[0].marker.style = XL_MARKER_STYLE.NONE
    if second_data:
        ch.series[1].format.line.color.rgb = _pal(1)
        ch.series[1].format.line.width = Pt(CHART_LINE_PT)
        ch.series[1].marker.style = XL_MARKER_STYLE.NONE

    if marker_idx is not None:
        ms = ch.series[marker_idx]
        ms.format.line.fill.background()
        ms.marker.style = XL_MARKER_STYLE.CIRCLE
        ms.marker.size = 9
        ms.marker.format.fill.solid()
        ms.marker.format.fill.fore_color.rgb = C_RED
        ms.marker.format.line.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        # 不用原生数据标签（会截断+黑引导线）；后面用文字框按精确坐标放

    # Y 轴自适应：漂亮整数刻度，范围包住数据最大/最小值
    allv = [d['close'] for d in data] + (
        [d['close'] for d in second_data] if second_data else [])
    ymin, ymax, ystep = _nice_axis(min(allv), max(allv))
    ch.value_axis.minimum_scale = ymin
    ch.value_axis.maximum_scale = ymax
    ch.value_axis.major_unit = ystep
    ch.value_axis.tick_labels.font.size = Pt(CHART_FONT_PT)

    # 关闭原生图例（改用手动图例，对齐曲线左侧）
    ch.has_legend = False
    ch.has_title = False

    style_chart_xml(ch, y_pct=y_pct)
    _set_plot_area(ch)           # 固定绘图区位置
    _set_chart_font(ch)          # 思源黑体

    # ── 黄色外框（无阴影）──
    border = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Cm(chart_l), Cm(chart_t), Cm(chart_w), Cm(chart_h))
    border.fill.background()
    border.line.color.rgb = C_YELLOW
    border.line.width = Pt(CHART_FRAME_PT)
    border.shadow.inherit = False

    # 绘图区精确几何（用于对齐图例/日期、精确放红点标签）
    plot_l = obj_l + PLOT_X * obj_w
    plot_t = obj_t + PLOT_Y * obj_h
    plot_w = PLOT_W * obj_w
    plot_h = PLOT_H * obj_h
    plot_r = plot_l + plot_w

    # ── 手动图例：橙短线 + 系列名，线与字垂直居中、左缘对齐曲线左侧 ──
    entries = [(series_name, _pal(0))]
    if second_data:
        entries.append((second_name, _pal(1)))
    lx, ly = plot_l, chart_t + FRAME_MARGIN
    for name, color in entries:
        sw = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                    Cm(lx), Cm(ly + LEG_H/2 - 0.05), Cm(0.8), Cm(0.10))
        sw.fill.solid(); sw.fill.fore_color.rgb = color
        sw.line.fill.background(); sw.shadow.inherit = False
        w_name = len(name) * (CHART_FONT_PT * 0.0353) + 0.3
        tb = add_tb(slide, name, lx + 1.0, ly, w_name + 0.6, LEG_H,
                    size=CHART_FONT_PT, bold=False, color=C_DARK, align=PP_ALIGN.LEFT)
        try:
            from pptx.enum.text import MSO_ANCHOR
            tb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        except Exception:
            pass
        lx += 1.0 + w_name + 0.8

    # ── 高低点标签：按精确坐标放在点旁，边缘避让，无引导线 ──
    denom = max(len(dates) - 1, 1)
    yspan = (ymax - ymin) or 1
    for i, p in marks:
        xr = i / denom
        yr = (p['close'] - ymin) / yspan
        x_cm = plot_l + xr * plot_w
        y_cm = plot_t + (1 - yr) * plot_h
        lw, lh = 4.6, 1.25
        if xr < 0.18:
            tx, al = x_cm + 0.3, PP_ALIGN.LEFT
        elif xr > 0.82:
            tx, al = x_cm - lw - 0.3, PP_ALIGN.RIGHT
        else:
            tx, al = x_cm - lw/2, PP_ALIGN.CENTER
        ty = (y_cm - lh - 0.2) if yr >= 0.5 else (y_cm + 0.25)   # 高点上方、低点下方留白
        label = f"{p['date']}\n收盘价{p['close']:,.2f}"
        add_tb(slide, label, tx, ty, lw, lh,
               size=13, bold=True, color=C_DARK, align=al)

    # ── X 轴起止日期：贴在曲线下方、对齐绘图区左右两端；与黄框底留白=FRAME_MARGIN ──
    date_y = chart_t + chart_h - FRAME_MARGIN - DATE_H
    add_tb(slide, data[0]['date'], plot_l, date_y, 6.5, DATE_H,
           size=16, bold=True, color=C_DARK, align=PP_ALIGN.LEFT)
    add_tb(slide, data[-1]['date'], plot_r - 6.5, date_y, 6.5, DATE_H,
           size=16, bold=True, color=C_DARK, align=PP_ALIGN.RIGHT)

    # ── 高亮区间（半透明背景色块 + 涨幅标注）──
    if bands:
        dn = max(len(dates) - 1, 1)
        def _xi(dstr):
            if dstr in dates: return dates.index(dstr)
            # 最近日期
            return min(range(len(dates)), key=lambda i: abs(
                (datetime.strptime(dates[i], '%Y-%m-%d') - datetime.strptime(dstr, '%Y-%m-%d')).days))
        for b in bands:
            i1, i2 = _xi(b['start']), _xi(b['end'])
            x1 = plot_l + (i1/dn)*plot_w
            x2 = plot_l + (i2/dn)*plot_w
            rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                          Cm(x1), Cm(plot_t), Cm(max(x2-x1, 0.2)), Cm(plot_h))
            _set_fill_alpha(rect, b.get('color', 'FF2700'), b.get('alpha', 22))
            rect.line.fill.background(); rect.shadow.inherit = False
            if b.get('label'):
                add_tb(slide, b['label'], x1 - 1.0, plot_t - 0.75, (x2-x1)+2.0, 0.7,
                       size=15, bold=True, color=C_RED, align=PP_ALIGN.CENTER)

    return cs_shape


# ── 高/低点标注 ───────────────────────────────────────────
def add_hilow(slide, data, hi_pt, lo_pt, chart_t, chart_h):
    dates  = [d['date'] for d in data]
    values = [d['close'] for d in data]
    v_min, v_max = min(values), max(values)
    v_rng = v_max - v_min or 1
    n = max(len(dates) - 1, 1)

    # plot 区估算（去掉图例和 Y 轴的占位）
    pl = CHART_L + 3.5    # 图例约占 3cm
    pr = CHART_L + CHART_W - 0.5
    pt_top = chart_t + 0.5
    pb = chart_t + chart_h - 1.3   # 底部日期区占 1.3cm

    for pt_data in [lo_pt, hi_pt]:
        if not pt_data:
            continue
        d_str = pt_data['date']
        v     = pt_data['close']
        if d_str not in dates:
            continue
        idx   = dates.index(d_str)
        x_rat = idx / n
        y_rat = (v - v_min) / v_rng

        x_cm  = pl + x_rat * (pr - pl)
        y_cm  = pb - y_rat * (pb - pt_top)

        # 红圆点
        dot = 0.32
        oval = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                      Cm(x_cm - dot/2), Cm(y_cm - dot/2),
                                      Cm(dot), Cm(dot))
        oval.fill.solid(); oval.fill.fore_color.rgb = C_RED
        oval.line.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); oval.line.width = Pt(1)
        oval.shadow.inherit = False

        dy = -1.4 if y_rat > 0.5 else 0.3
        dx = -4.0 if x_rat > 0.65 else 0.2

        label = f'{d_str}\n收盘价{v:.2f}'
        tx = max(CHART_L + 0.1, min(x_cm + dx, CHART_L + CHART_W - 4.2))
        ty = max(chart_t + 0.1, min(y_cm + dy, chart_t + chart_h - 1.8))
        add_tb(slide, label, tx, ty, 4.0, 1.0,
               size=12, bold=True, color=C_DARK)


# ── 标准正文图表页 ────────────────────────────────────────
def chart_page(prs, layout, slide_title, data,
               points=None, extra_hi=None, extra_lo=None,
               second_data=None, y_pct=False, show_hilo=True, bands=None,
               series_name='中证全指', second_name='中证全指全收益',
               source='数据来源：Wind'):
    hi, lo = find_hi_lo(data)
    dm = {d['date']: d for d in data}
    hi_pt = dm.get(extra_hi, hi) if extra_hi else hi
    lo_pt = dm.get(extra_lo, lo) if extra_lo else lo

    sl = prs.slides.add_slide(layout)

    # 清空 placeholder（避免默认文字框干扰）
    for ph in list(sl.placeholders):
        ph.element.getparent().remove(ph.element)

    # ── 标题（40pt，黑色，居中）──
    t_pt = title_fit(slide_title)
    t_h  = text_height_cm(slide_title, t_pt, CONTENT_W) + 0.2
    add_tb(sl, slide_title,
           CONTENT_L, CONTENT_TOP, CONTENT_W, t_h,
           size=t_pt, bold=True, color=C_BLACK, align=PP_ALIGN.CENTER)

    y_top = CONTENT_TOP + t_h + 0.3
    MAX_RATIO = 2.5
    text_heavy = bool(points) and len(points) >= 2

    if text_heavy:
        # ── 左右排版：文字左、图右（图更大）──
        LEFT_W = 12.5
        add_arrow_points(sl, points, y_top, left=CONTENT_L, width=LEFT_W - 0.3)
        region_l = CONTENT_L + LEFT_W
        region_w = CONTENT_W - LEFT_W
        chart_t = y_top
        chart_h = CONTENT_BOTTOM - 0.6 - chart_t
        chart_w = region_w
        if chart_w / chart_h > MAX_RATIO:
            chart_w = chart_h * MAX_RATIO
        chart_l = region_l + (region_w - chart_w) / 2
    else:
        # ── 上下排版：标题(+少量要点) 在上，图在下 ──
        y_cursor = y_top
        if points:
            y_cursor = add_arrow_points(sl, points, y_cursor) + 0.2
        chart_t = y_cursor
        chart_h = CONTENT_BOTTOM - 0.6 - chart_t
        chart_w = CONTENT_W
        if chart_w / chart_h > MAX_RATIO:
            chart_w = chart_h * MAX_RATIO
        chart_l = CONTENT_L + (CONTENT_W - chart_w) / 2

    add_line_chart(sl, data, chart_l, chart_t, chart_w, chart_h,
                   second_data=second_data, y_pct=y_pct,
                   hi_pt=(hi_pt if show_hilo else None),
                   lo_pt=(lo_pt if show_hilo else None), bands=bands,
                   series_name=series_name, second_name=second_name)

    # 数据来源（黄框右下外侧）
    add_tb(sl, source,
           chart_l + chart_w - 6.5, chart_t + chart_h + 0.05, 6.5, 0.6,
           size=12, color=C_GRAY, align=PP_ALIGN.RIGHT)

    return sl


# ── 封面页（规范E1：自适应+末行压横线）──────────────────
COVER_L, COVER_W, COVER_LINE_Y = 1.71, 17.75, 7.80  # 母版标题占位底部=横线

def _split_lines(text, nlines):
    """把 text 切成 nlines 行：在最接近等分点的标点后断行，避免拆词。"""
    puncts = '：，。、；？！）」'
    n = len(text)
    if nlines <= 1:
        return [text]
    lines, start = [], 0
    for k in range(1, nlines):
        ideal = round(n * k / nlines)
        # 在 ideal 附近找最近的标点（断在标点后）
        best, bd = None, 999
        for i in range(start, n):
            if text[i] in puncts and abs(i - ideal) < bd:
                best, bd = i, abs(i - ideal)
        cut = (best + 1) if best is not None else max(start + 1, ideal)
        lines.append(text[start:cut]); start = cut
    lines.append(text[start:])
    return [s for s in lines if s]

def cover_page(prs, layout, title_text):
    sl = prs.slides.add_slide(layout)
    for ph in list(sl.placeholders):
        ph.element.getparent().remove(ph.element)

    # 自适应：选行数让最长行能在宽度内、字号尽量大（40~60）。
    # 如果标题已手动换行，尊重换行，避免逗号/单字被算法单独挤到一行。
    if '\n' in title_text:
        lines = [s for s in title_text.split('\n') if s]
        longest = max(len(s) for s in lines)
        pt = max(36, min(COVER_PT, int(COVER_W / (longest * 0.0353))))
    else:
        best = None
        for nlines in (1, 2, 3, 4):
            lines = _split_lines(title_text, nlines)
            longest = max(len(s) for s in lines)
            pt = min(COVER_PT, int(COVER_W / (longest * 0.0353)))
            if pt >= 40 or nlines == 4:
                best = (max(pt, 36), lines)
                break
        pt, lines = best
    wrapped = '\n'.join(lines)

    block_h = line_height_cm(pt) * len(lines)
    top = COVER_LINE_Y - block_h        # 末行底压横线
    add_tb(sl, wrapped, COVER_L, top, COVER_W, block_h + 0.6,
           size=pt, bold=True, color=C_MAIN, align=PP_ALIGN.LEFT)
    return sl


# ── 纯文字页 ──────────────────────────────────────────────
def text_page(prs, layout, slide_title, body_text):
    sl = prs.slides.add_slide(layout)
    for ph in list(sl.placeholders):
        ph.element.getparent().remove(ph.element)

    t_pt = title_fit(slide_title)
    t_h  = text_height_cm(slide_title, t_pt, CONTENT_W) + 0.2
    add_tb(sl, slide_title,
           CONTENT_L, CONTENT_TOP, CONTENT_W, t_h,
           size=t_pt, bold=True, color=C_BLACK, align=PP_ALIGN.CENTER)

    y = CONTENT_TOP + t_h + 0.3
    b_pt = auto_pt(body_text, BODY_PT)
    b_h  = CONTENT_BOTTOM - y - 0.5
    add_tb(sl, body_text,
           CONTENT_L, y, CONTENT_W, b_h,
           size=b_pt, bold=False, color=C_BLACK, align=PP_ALIGN.LEFT)
    return sl


# ── 橙色箭头要点（规范 B5.1：▶ + 黑色正文，关键词加粗）──────
def add_arrow_points(sl, points, top, pt=BODY_PT, left=CONTENT_L, width=CONTENT_W):
    """points: list[str]，逐条加 ▶ 要点。返回结束 y。"""
    y = top
    for txt in points:
        h = text_height_cm(txt, pt, width - 0.9) + 0.15
        # 箭头
        add_tb(sl, '▶', left, y, 0.8, line_height_cm(pt) + 0.1,
               size=pt - 4, bold=True, color=C_MAIN, align=PP_ALIGN.LEFT)
        # 文字
        add_tb(sl, txt, left + 0.9, y, width - 0.9, h,
               size=pt, bold=False, color=C_DARK, align=PP_ALIGN.LEFT)
        y += h + 0.1
    return y


# ── 图片页（规范：标题居中 + 大图，走势图自带黄框不再加框）──
def img_page(prs, layout, slide_title, img_path,
             points=None, pending=False, big=True):
    """
    标题居中(主色) + 可选橙箭头要点 + 居中大图。
    pending=True 时右下角加灰字「⚠️图片待更新」(规范 B3.2)。
    big=True(默认): 纯图页尽量放大——压缩标题与图间距、向下用满到接近页脚、
                    宽度允许铺到整页(信息图看清为先)。
    """
    sl = prs.slides.add_slide(layout)
    for ph in list(sl.placeholders):
        ph.element.getparent().remove(ph.element)

    # 标题
    t_pt = title_fit(slide_title)
    t_h  = text_height_cm(slide_title, t_pt, CONTENT_W) + 0.2
    add_tb(sl, slide_title, CONTENT_L, CONTENT_TOP, CONTENT_W, t_h,
           size=t_pt, bold=True, color=C_BLACK, align=PP_ALIGN.CENTER)
    y = CONTENT_TOP + t_h + (0.1 if big else 0.25)

    # 箭头要点
    if points:
        y = add_arrow_points(sl, points, y) + 0.2

    # 图片：等比缩放，放进剩余区域并水平居中
    # big 模式向下用到 18.3(贴近页脚)、宽度铺到整页(33.4)，把信息图放到最大
    if big and not points:
        avail_h = 18.3 - y
        avail_w = 33.4
        L0 = 0.25
    else:
        avail_h = CONTENT_BOTTOM - y - 0.4
        avail_w = CONTENT_W
        L0 = CONTENT_L
    iw, ih = Image.open(img_path).size
    ratio = min(avail_w / (iw / 96 * 2.54), avail_h / (ih / 96 * 2.54))
    # 用纯比例缩放：以可用宽高为界
    w_cm = iw * 0.02646   # px→cm @96dpi
    h_cm = ih * 0.02646
    scale = min(avail_w / w_cm, avail_h / h_cm)
    w_cm *= scale; h_cm *= scale
    left = L0 + (avail_w - w_cm) / 2
    sl.shapes.add_picture(img_path, Cm(left), Cm(y), Cm(w_cm), Cm(h_cm))

    if pending:
        add_tb(sl, '⚠️图片待更新',
               CHART_L + CHART_W - 5.0, CONTENT_BOTTOM - 0.6, 5.0, 0.6,
               size=14, color=C_GRAY, align=PP_ALIGN.RIGHT)
    return sl


# ── 三个周期可视化（规范：三个框 + 气泡）────────────────────
def box_page(prs, layout, slide_title, boxes, bubble=None):
    """N 个圆角框（标题+说明）+ 可选气泡。boxes: [(标题, 说明), ...]"""
    sl = prs.slides.add_slide(layout)
    for ph in list(sl.placeholders):
        ph.element.getparent().remove(ph.element)
    t_pt = title_fit(slide_title)
    t_h  = text_height_cm(slide_title, t_pt, CONTENT_W) + 0.2
    add_tb(sl, slide_title, CONTENT_L, CONTENT_TOP, CONTENT_W, t_h,
           size=t_pt, bold=True, color=C_BLACK, align=PP_ALIGN.CENTER)

    n = len(boxes)
    box_top = CONTENT_TOP + t_h + 1.2
    gap = 1.0
    side = 1.0                               # G: 两侧留边，不贴边
    box_w = (CONTENT_W - 2*side - (n - 1) * gap) / n
    ttl_pt = 30 if n <= 3 else 24
    desc_pt = 18                             # 缩小正文，避免溢出
    # F/K: 盒高按最长说明自适应（标题区1.4 + 间隔0.5 + 文字行高 + 上下内边距）
    desc_lines = max(estimate_lines(d, desc_pt, box_w - 1.0) for _, d in boxes)
    box_h = 1.4 + 0.5 + desc_lines * line_height_cm(desc_pt) + 1.0
    tints = [RGBColor(0xFC,0xE6,0xD2), RGBColor(0xFC,0xE6,0xD2)]   # 素材库浅橙底
    for i, (ttl, desc) in enumerate(boxes):
        x = CONTENT_L + side + i * (box_w + gap)
        box = sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  Cm(x), Cm(box_top), Cm(box_w), Cm(box_h))
        box.fill.solid(); box.fill.fore_color.rgb = tints[i % 2]
        box.line.color.rgb = C_MAIN; box.line.width = Pt(2.0)
        box.shadow.inherit = False
        box.text_frame.text = ''
        add_tb(sl, ttl, x + 0.5, box_top + 0.6, box_w - 1.0, 1.4,
               size=ttl_pt, bold=True, color=C_MAIN, align=PP_ALIGN.CENTER)
        add_tb(sl, desc, x + 0.5, box_top + 2.1, box_w - 1.0, box_h - 2.5,
               size=desc_pt, bold=False, color=C_DARK, align=PP_ALIGN.LEFT)

    if bubble:
        bub_top = box_top + box_h + 0.9
        bub_lines = estimate_lines(bubble, 20, CONTENT_W - 6)
        bub_h = max(2.0, bub_lines * line_height_cm(20) + 0.9)
        bub = sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  Cm(CONTENT_L + 2), Cm(bub_top), Cm(CONTENT_W - 4), Cm(bub_h))
        bub.fill.solid(); bub.fill.fore_color.rgb = RGBColor(0xFC,0xE6,0xD2)
        bub.line.fill.background(); bub.shadow.inherit = False
        bub.text_frame.text = ''
        tb = add_tb(sl, bubble, CONTENT_L + 2.5, bub_top, CONTENT_W - 5, bub_h,
               size=20, bold=False, color=C_DARK, align=PP_ALIGN.CENTER)
        tb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    return sl

def three_cycle_page(prs, layout, slide_title, cycles, bubble):
    """兼容旧接口：N框+气泡。"""
    return box_page(prs, layout, slide_title, cycles, bubble)


# ── 6轮牛熊汇总表（抓数重建为原生表格，规范B4表头橙底白字）──
def summary_table_page(prs, layout, title, phases):
    """phases: [(轮次, 类型, start, end), ...] 抓Wind算持续天数+最大涨跌幅"""
    from datetime import datetime
    sl = prs.slides.add_slide(layout)
    for ph in list(sl.placeholders):
        ph.element.getparent().remove(ph.element)

    add_tb(sl, title, CONTENT_L, CONTENT_TOP, CONTENT_W,
           text_height_cm(title, TITLE_PT, CONTENT_W) + 0.2,
           size=TITLE_PT, bold=True, color=C_BLACK, align=PP_ALIGN.CENTER)

    headers = ['轮次', '牛/熊', '起止时间', '持续天数', '最大涨跌幅']
    rows = []
    for rn, typ, start, end in phases:
        print(f'  表格取数 {rn}{typ}: {start}~{end}')
        d = fetch_kline(WINDCODE, start, end)
        closes = [x['close'] for x in d]
        mn, mx = min(closes), max(closes)
        chg = (mx/mn - 1)*100 if typ == '牛' else (mn/mx - 1)*100
        days = (datetime.strptime(end, '%Y-%m-%d') - datetime.strptime(start, '%Y-%m-%d')).days
        rows.append([rn, typ, f'{start} ~ {end}', f'{days}', f'{chg:+.2f}%'])

    nr, nc = len(rows) + 1, len(headers)
    top = CONTENT_TOP + 1.4
    tbl_h = CONTENT_BOTTOM - top - 0.3
    gtbl = sl.shapes.add_table(nr, nc, Cm(CONTENT_L), Cm(top),
                               Cm(CONTENT_W), Cm(tbl_h)).table
    # 第一列加宽，让"第一轮"等不换行；总宽=CONTENT_W
    widths = [3.6, 2.6, 13.5, 5.0, 7.3]
    for c, w in enumerate(widths):
        gtbl.columns[c].width = Cm(w)
    # 固定行高，避免整表过高/溢出
    row_h = min(1.05, tbl_h / nr)
    for r in range(nr):
        gtbl.rows[r].height = Cm(row_h)

    def cell(r, c, text, bold=False, color=C_DARK, fill=None, align=PP_ALIGN.CENTER):
        cl = gtbl.cell(r, c)
        cl.fill.solid()
        cl.fill.fore_color.rgb = fill if fill is not None else RGBColor(0xFF,0xFF,0xFF)
        cl.margin_top = Cm(0.02); cl.margin_bottom = Cm(0.02)
        cl.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = cl.text_frame; tf.word_wrap = False
        p = tf.paragraphs[0]; p.alignment = align
        run = p.add_run(); run.text = text
        run.font.size = Pt(18); run.font.bold = bold
        run.font.color.rgb = color; set_run_font(run)

    for c, h in enumerate(headers):
        cell(0, c, h, bold=True, color=RGBColor(0xFF,0xFF,0xFF), fill=C_MAIN)
    for r, row in enumerate(rows, start=1):
        bg = RGBColor(0xF5,0xF5,0xF5) if r % 2 == 0 else RGBColor(0xFF,0xFF,0xFF)
        for c, val in enumerate(row):
            cell(r, c, val, bold=(c == 0),
                 color=(C_MAIN if c == 4 else C_DARK), fill=bg,
                 align=(PP_ALIGN.LEFT if c == 2 else PP_ALIGN.CENTER))
    add_tb(sl, '数据来源：Wind', CONTENT_L + CONTENT_W - 6.5,
           CONTENT_BOTTOM - 0.6, 6.5, 0.6, size=14, color=C_GRAY,
           align=PP_ALIGN.RIGHT)
    return sl




# ══════════════════════════════════════════════════════════
#  通用管线：根据"内容spec"生成整套PPT（引擎部分到此对外）
# ══════════════════════════════════════════════════════════
def arrows_page(prs, layout, title, points, pt=22, spread=False):
    """标题居中 + 橙色箭头要点（总结/纯要点页）。spread=True 时纵向均匀铺满版面(J:不留大片白)。"""
    sl = prs.slides.add_slide(layout)
    for ph in list(sl.placeholders):
        ph.element.getparent().remove(ph.element)
    t_h = text_height_cm(title, TITLE_PT, CONTENT_W) + 0.2
    add_tb(sl, title, CONTENT_L, CONTENT_TOP, CONTENT_W, t_h,
           size=TITLE_PT, bold=True, color=C_BLACK, align=PP_ALIGN.CENTER)
    start = CONTENT_TOP + t_h + 1.0
    if spread:
        avail = CONTENT_BOTTOM - start
        slot = avail / len(points)
        for i, txt in enumerate(points):
            y = start + i * slot
            h = text_height_cm(txt, pt, CONTENT_W - 0.9) + 0.15
            add_tb(sl, '▶', CONTENT_L, y + (slot - h)/2 - 0.05, 0.9, line_height_cm(pt)+0.1,
                   size=pt-4, bold=True, color=C_MAIN, align=PP_ALIGN.LEFT)
            add_tb(sl, txt, CONTENT_L + 1.0, y + (slot - h)/2, CONTENT_W - 1.0, h,
                   size=pt, bold=False, color=C_DARK, align=PP_ALIGN.LEFT)
    else:
        add_arrow_points(sl, points, start, pt=pt)
    return sl


def _fill_points(points, data):
    """把要点里的 {lo}{hi}{rise}{hi_month}{lo_date}{hi_date} 用数据填上。"""
    if not points:
        return points
    lo = min(data, key=lambda x: x['close'])
    hi = max(data, key=lambda x: x['close'])
    rise = (hi['close'] - lo['close']) / lo['close'] * 100
    ctx = {'lo': f"{lo['close']:.2f}", 'hi': f"{hi['close']:.2f}",
           'rise': f"{rise:.2f}", 'hi_month': hi['date'][:7],
           'lo_date': lo['date'], 'hi_date': hi['date']}
    out = []
    for p in points:
        out.append(p.format(**ctx) if '{' in p else p)
    return out


def build_deck(pages, out_path, tmpl=TMPL_PATH,
               windcode=WINDCODE, img_dir=IMG_DIR, drop_tmpl_slides=8):
    """
    根据 pages（每页一个 dict）生成整套 PPT。页type:
      cover / table / chart / chart_dual / image / three_cycle / summary
    """
    import os
    prs = Presentation(tmpl)
    body  = get_layout(prs, 1, '自定义版式')
    cover = get_layout(prs, 4, '标题幻灯片')

    for pg in pages:
        t = pg.get('type')
        print(f'[{t}] {pg.get("title","")[:20]}')
        if t == 'cover':
            cover_page(prs, cover, pg['title'])
        elif t == 'section':
            section_page(prs, pg['title'], pg.get('number', ''))
        elif t == 'table':
            summary_table_page(prs, body, pg['title'], pg['phases'])
        elif t == 'chart':
            # 手动数据：spec 直接给 data=[{date,close},...]（外汇等Wind取不到的页）
            if pg.get('data'):
                d = pg['data']
            else:
                d = fetch_kline(pg.get('windcode', windcode), pg['start'], pg['end'],
                                pg.get('freq', 'D'))
            pts = _fill_points(pg.get('points'), d)
            hi = lo = None
            if pg.get('auto_hilo'):
                hi = max(d, key=lambda x: x['close'])['date']
                lo = min(d, key=lambda x: x['close'])['date']
            chart_page(prs, body, pg['title'], d, points=pts,
                       extra_hi=pg.get('hi', hi), extra_lo=pg.get('lo', lo),
                       show_hilo=pg.get('show_hilo', True), bands=pg.get('bands'))
        elif t == 'chart_dual':
            fq = pg.get('freq', 'D')
            dm = fetch_kline(pg.get('base_code', windcode), pg['start'], pg['end'], fq)
            dt = fetch_kline(pg['second_code'], pg['start'], pg['end'], fq)
            bm, bt = dm[0]['close'], dt[0]['close']
            pm = [{'date': x['date'], 'close': (x['close']/bm-1)*100} for x in dm]
            pt = [{'date': x['date'], 'close': (x['close']/bt-1)*100} for x in dt]
            chart_page(prs, body, pg['title'], pm, points=pg.get('points'),
                       second_data=pt, y_pct=True, show_hilo=False,
                       series_name=pg.get('series_name', '中证全指'),
                       second_name=pg.get('second_name', '中证全指全收益'))
        elif t == 'chart_multi':
            fq = pg.get('freq', 'D')
            mode = pg.get('mode', 'percent')
            sl_list = []
            for s in pg['series']:
                if s.get('data'):
                    d = s['data']
                else:
                    d = fetch_kline(s.get('code', windcode), pg['start'], pg['end'], fq)
                if mode == 'percent':
                    base = d[0]['close']
                    vals = [(x['close']/base - 1)*100 for x in d]
                else:
                    vals = [x['close'] for x in d]
                sl_list.append({'name': s['name'], 'dates': [x['date'] for x in d], 'values': vals})
            if pg.get('left_table'):
                multi_chart_table_page(prs, body, pg['title'], sl_list,
                                       table_title=pg.get('table_title', '区间最大涨幅'),
                                       y_pct=(mode == 'percent'))
            else:
                multi_chart_page(prs, body, pg['title'], sl_list,
                                 points=pg.get('points'), y_pct=(mode == 'percent'))
        elif t == 'chart_below_table':
            fq = pg.get('freq', 'D'); mode = pg.get('mode', 'percent')
            sl_list = []
            for s in pg['series']:
                d = s['data'] if s.get('data') else fetch_kline(s.get('code', windcode), pg['start'], pg['end'], fq)
                base = d[0]['close']
                vals = [(x['close']/base-1)*100 for x in d] if mode == 'percent' else [x['close'] for x in d]
                sl_list.append({'name': s['name'], 'dates': [x['date'] for x in d], 'values': vals})
            chart_below_table_page(prs, body, pg['title'], sl_list,
                                   pg['table']['headers'], pg['table']['rows'],
                                   y_pct=(mode == 'percent'))
        elif t == 'image':
            img = pg['img']
            if not os.path.isabs(img):
                img = os.path.join(img_dir, img)
            img_page(prs, body, pg['title'], img,
                     points=pg.get('points'), pending=pg.get('pending', False))
        elif t in ('three_cycle', 'boxes'):
            box_page(prs, body, pg['title'],
                     [tuple(c) for c in pg.get('boxes', pg.get('cycles'))], pg.get('bubble'))
        elif t in ('summary', 'arrows'):
            arrows_page(prs, body, pg['title'], pg['points'], pt=pg.get('pt', 22))
        else:
            print(f'  ⚠ 未知页type: {t}')

    for _ in range(drop_tmpl_slides):
        remove_slide(prs, 0)
    prs.save(out_path)
    print(f'\n✅ 完成！{len(prs.slides)} 张 → {out_path}')
    return prs


# ── 过渡页 / 标题页（只有标题、无正文/图的页用这个模版）──────
def get_section_layout(prs):
    """找母版里的过渡页版式（含'02'编号占位 + '单击此处'文本框）。"""
    for master in prs.slide_masters:
        for lyt in master.slide_layouts:
            xml = lyt._element.xml
            if '单击此处' in xml and '单击此处编辑母版' not in xml:
                return lyt
    return None

def section_page(prs, title, number=''):
    """过渡页：橙色编号 + 大标题。number 可空。"""
    lyt = get_section_layout(prs)
    if lyt is None:                       # 兜底：没有过渡页版式就用正文版式居中大标题
        sl = prs.slides.add_slide(get_layout(prs, 1, '自定义版式'))
        for ph in list(sl.placeholders):
            ph.element.getparent().remove(ph.element)
        add_tb(sl, title, CONTENT_L, 7.0, CONTENT_W, 3.0,
               size=54, bold=True, color=C_MAIN, align=PP_ALIGN.CENTER)
        return sl
    sl = prs.slides.add_slide(lyt)
    # title占位符(原'02'位置) → 编号
    for ph in list(sl.placeholders):
        if ph.placeholder_format.type is not None and 'TITLE' in str(ph.placeholder_format.type):
            ph.text = str(number)
            for para in ph.text_frame.paragraphs:
                for run in para.runs:
                    run.font.bold = True; run.font.color.rgb = C_MAIN
                    set_run_font(run)
            break
    # 在“单击此处”位置(≈5.4,8.8)叠白底文本框放标题
    tb = add_tb(sl, title, 5.0, 8.5, 17.5, 2.2,
                size=40, bold=True, color=C_BLACK, align=PP_ALIGN.LEFT)
    tb.fill.solid(); tb.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    tb.line.fill.background()
    return sl


# ── 多线走势图（N条线，复用样式；不标红点）─────────────────
def add_multi_line_chart(slide, chart_l, chart_t, chart_w, chart_h, series_list, y_pct=True):
    SIDE = 0.7; LEG_H, DATE_H = 0.70, 0.65
    TOP_BAND = FRAME_MARGIN + LEG_H
    BOT_BAND = FRAME_MARGIN + DATE_H + 0.15
    obj_l, obj_t = chart_l + SIDE, chart_t + TOP_BAND
    obj_w, obj_h = chart_w - 2*SIDE, chart_h - TOP_BAND - BOT_BAND

    # X轴用所有线共同的交易日(交集)，保证每条线每点都有值，绝不断线
    common = set(series_list[0]['dates'])
    for s in series_list[1:]:
        common &= set(s['dates'])
    cats = [d for d in series_list[0]['dates'] if d in common]
    cd = ChartData(); cd.categories = cats
    for s in series_list:
        dm = dict(zip(s['dates'], s['values']))
        cd.add_series(s['name'], [dm[d] for d in cats])
    cs = slide.shapes.add_chart(XL_CHART_TYPE.LINE, Cm(obj_l), Cm(obj_t), Cm(obj_w), Cm(obj_h), cd)
    ch = cs.chart
    for i, s in enumerate(series_list):
        ch.series[i].format.line.color.rgb = _pal(i)
        ch.series[i].format.line.width = Pt(CHART_LINE_PT)
        ch.series[i].marker.style = XL_MARKER_STYLE.NONE

    allv = [v for s in series_list for v in s['values'] if v is not None]
    ymin, ymax, ystep = _nice_axis(min(allv), max(allv))
    ch.value_axis.minimum_scale = ymin; ch.value_axis.maximum_scale = ymax
    ch.value_axis.major_unit = ystep
    ch.value_axis.tick_labels.font.size = Pt(CHART_FONT_PT)
    ch.has_legend = False; ch.has_title = False
    style_chart_xml(ch, y_pct=y_pct); _set_plot_area(ch); _set_chart_font(ch)

    border = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(chart_l), Cm(chart_t), Cm(chart_w), Cm(chart_h))
    border.fill.background(); border.line.color.rgb = C_YELLOW
    border.line.width = Pt(CHART_FRAME_PT); border.shadow.inherit = False

    plot_l = obj_l + PLOT_X * obj_w
    plot_r = plot_l + PLOT_W * obj_w
    # 手动图例（多条，超宽则换行）；字号按条数收缩
    fs = CHART_FONT_PT if len(series_list) <= 3 else 15
    lx, ly = plot_l, chart_t + FRAME_MARGIN
    for i, s in enumerate(series_list):
        w_name = len(s['name']) * (fs * 0.0353) + 0.3
        if lx + 1.0 + w_name > plot_r:        # 换行
            lx = plot_l; ly += LEG_H
        sw = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                    Cm(lx), Cm(ly + LEG_H/2 - 0.05), Cm(0.7), Cm(0.10))
        sw.fill.solid(); sw.fill.fore_color.rgb = _pal(i)
        sw.line.fill.background(); sw.shadow.inherit = False
        tb = add_tb(slide, s['name'], lx + 0.9, ly, w_name + 0.6, LEG_H,
                    size=fs, bold=False, color=C_DARK, align=PP_ALIGN.LEFT)
        try: tb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        except Exception: pass
        lx += 0.9 + w_name + 0.6

    date_y = chart_t + chart_h - FRAME_MARGIN - DATE_H
    add_tb(slide, cats[0], plot_l, date_y, 6.5, DATE_H, size=16, bold=True, color=C_DARK, align=PP_ALIGN.LEFT)
    add_tb(slide, cats[-1], plot_r - 6.5, date_y, 6.5, DATE_H, size=16, bold=True, color=C_DARK, align=PP_ALIGN.RIGHT)
    return cs


def multi_chart_page(prs, layout, title, series_list, points=None, y_pct=True):
    sl = prs.slides.add_slide(layout)
    for ph in list(sl.placeholders):
        ph.element.getparent().remove(ph.element)
    t_pt = title_fit(title)
    t_h = text_height_cm(title, t_pt, CONTENT_W) + 0.2
    add_tb(sl, title, CONTENT_L, CONTENT_TOP, CONTENT_W, t_h,
           size=t_pt, bold=True, color=C_BLACK, align=PP_ALIGN.CENTER)
    y_top = CONTENT_TOP + t_h + 0.3
    MAX_RATIO = 2.5
    if points and len(points) >= 2:
        LEFT_W = 12.5
        add_arrow_points(sl, points, y_top, left=CONTENT_L, width=LEFT_W - 0.3)
        region_l, region_w = CONTENT_L + LEFT_W, CONTENT_W - LEFT_W
        ct = y_top; chh = CONTENT_BOTTOM - 0.6 - ct
        cw = region_w
        if cw / chh > MAX_RATIO: cw = chh * MAX_RATIO
        cl = region_l + (region_w - cw) / 2
    else:
        y = y_top
        if points: y = add_arrow_points(sl, points, y) + 0.2
        ct = y; chh = CONTENT_BOTTOM - 0.6 - ct
        cw = CONTENT_W
        if cw / chh > MAX_RATIO: cw = chh * MAX_RATIO
        cl = CONTENT_L + (CONTENT_W - cw) / 2
    add_multi_line_chart(sl, cl, ct, cw, chh, series_list, y_pct=y_pct)
    add_tb(sl, '数据来源：Wind', cl + cw - 6.5, ct + chh + 0.05, 6.5, 0.6,
           size=12, color=C_GRAY, align=PP_ALIGN.RIGHT)
    return sl


# ── 多线图 + 左侧表格（区间最大涨幅）──────────────────────
def multi_chart_table_page(prs, layout, title, series_list, table_title='累计涨幅',
                           y_pct=True, points=None):
    sl = prs.slides.add_slide(layout)
    for ph in list(sl.placeholders):
        ph.element.getparent().remove(ph.element)
    t_pt = title_fit(title)
    t_h = text_height_cm(title, t_pt, CONTENT_W) + 0.2
    add_tb(sl, title, CONTENT_L, CONTENT_TOP, CONTENT_W, t_h,
           size=t_pt, bold=True, color=C_BLACK, align=PP_ALIGN.CENTER)
    y_top = CONTENT_TOP + t_h + 0.4
    if points:
        y_top = add_arrow_points(sl, points, y_top) + 0.2

    # 左侧表格：指数 | 区间最大涨幅（从各线累计%取峰值）
    TBL_W = 8.5
    rows = [(s['name'].replace('全收益', ''), f"{s['values'][-1]:+.2f}%") for s in series_list]
    nr = len(rows) + 1
    tbl_h = min(1.0 * nr, CONTENT_BOTTOM - y_top - 0.5)
    tbl_top = y_top + 1.5
    gtbl = sl.shapes.add_table(nr, 2, Cm(CONTENT_L), Cm(tbl_top), Cm(TBL_W), Cm(tbl_h)).table
    gtbl.columns[0].width = Cm(5.0); gtbl.columns[1].width = Cm(3.5)
    for r in range(nr):
        gtbl.rows[r].height = Cm(tbl_h / nr)
    def cell(r, c, text, bold=False, color=C_DARK, fill=None):
        cl = gtbl.cell(r, c)
        cl.fill.solid(); cl.fill.fore_color.rgb = fill or RGBColor(0xFF, 0xFF, 0xFF)
        cl.margin_top = Cm(0.02); cl.margin_bottom = Cm(0.02)
        cl.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = cl.text_frame; tf.word_wrap = False
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        run = p.add_run(); run.text = text
        run.font.size = Pt(16); run.font.bold = bold; run.font.color.rgb = color
        set_run_font(run)
    cell(0, 0, '指数', True, RGBColor(0xFF,0xFF,0xFF), C_MAIN)
    cell(0, 1, table_title, True, RGBColor(0xFF,0xFF,0xFF), C_MAIN)
    for i, (nm, val) in enumerate(rows, 1):
        bg = RGBColor(0xF5,0xF5,0xF5) if i % 2 == 0 else RGBColor(0xFF,0xFF,0xFF)
        cell(i, 0, nm, True, C_DARK, bg)
        cell(i, 1, val, False, C_MAIN, bg)

    # 右侧多线图
    region_l = CONTENT_L + TBL_W + 0.6
    region_w = CONTENT_W - TBL_W - 0.6
    ct = y_top
    chh = CONTENT_BOTTOM - 0.6 - ct
    cw = region_w
    if cw / chh > 2.5: cw = chh * 2.5
    cl = region_l + (region_w - cw) / 2
    add_multi_line_chart(sl, cl, ct, cw, chh, series_list, y_pct=y_pct)
    add_tb(sl, '数据来源：Wind', cl + cw - 6.5, ct + chh + 0.05, 6.5, 0.6,
           size=12, color=C_GRAY, align=PP_ALIGN.RIGHT)
    return sl


# ── 多线图 + 下方数据表格（如 P5 盈利同比）──────────────────
def chart_below_table_page(prs, layout, title, series_list,
                           table_headers, table_rows, y_pct=True, points=None):
    sl = prs.slides.add_slide(layout)
    for ph in list(sl.placeholders):
        ph.element.getparent().remove(ph.element)
    t_pt = title_fit(title)
    t_h = text_height_cm(title, t_pt, CONTENT_W) + 0.2
    add_tb(sl, title, CONTENT_L, CONTENT_TOP, CONTENT_W, t_h,
           size=t_pt, bold=True, color=C_BLACK, align=PP_ALIGN.CENTER)
    y_top = CONTENT_TOP + t_h + 0.3
    if points:
        y_top = add_arrow_points(sl, points, y_top) + 0.2

    # 表格高度（按行数），图占其余
    nr = len(table_rows) + 1
    tbl_h = min(0.95 * nr, 4.5)
    tbl_top = CONTENT_BOTTOM - tbl_h
    # 图：从 y_top 到表格上方
    ch_h = tbl_top - 0.5 - y_top
    add_multi_line_chart(sl, CONTENT_L, y_top, CONTENT_W, ch_h, series_list, y_pct=y_pct)

    # 下方表格
    nc = len(table_headers)
    gtbl = sl.shapes.add_table(nr, nc, Cm(CONTENT_L), Cm(tbl_top),
                               Cm(CONTENT_W), Cm(tbl_h)).table
    col0 = 5.5
    gtbl.columns[0].width = Cm(col0)
    for c in range(1, nc):
        gtbl.columns[c].width = Cm((CONTENT_W - col0) / (nc - 1))
    for r in range(nr):
        gtbl.rows[r].height = Cm(tbl_h / nr)
    def cell(r, c, text, bold=False, color=C_DARK, fill=None, align=PP_ALIGN.CENTER):
        cl = gtbl.cell(r, c)
        cl.fill.solid(); cl.fill.fore_color.rgb = fill or RGBColor(0xFF,0xFF,0xFF)
        cl.margin_top = Cm(0.02); cl.margin_bottom = Cm(0.02)
        cl.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = cl.text_frame; tf.word_wrap = False
        p = tf.paragraphs[0]; p.alignment = align
        run = p.add_run(); run.text = str(text)
        run.font.size = Pt(15); run.font.bold = bold; run.font.color.rgb = color
        set_run_font(run)
    for c, hdr in enumerate(table_headers):
        cell(0, c, hdr, True, RGBColor(0xFF,0xFF,0xFF), C_MAIN)
    for r, row in enumerate(table_rows, 1):
        bg = RGBColor(0xF5,0xF5,0xF5) if r % 2 == 0 else RGBColor(0xFF,0xFF,0xFF)
        for c, val in enumerate(row):
            cell(r, c, val, bold=(c == 0), color=C_DARK, fill=bg,
                 align=(PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER))
    add_tb(sl, '数据来源：Wind，螺丝钉团队', CONTENT_L + CONTENT_W - 8, tbl_top - 0.6, 8, 0.5,
           size=11, color=C_GRAY, align=PP_ALIGN.RIGHT)
    return sl


# ── 原生双轴组合图：主轴线 + 次轴面积 + 次轴标记点（P4型，可编辑）──
def add_dual_axis_chart(slide, chart_l, chart_t, chart_w, chart_h,
                        dates, line_name, line_vals, area_name, area_vals,
                        marker_series=None, source='数据来源：Wind'):
    """美元指数=橙线(左轴) + 国债收益率=面积(右轴) + 加息/降息点(打在美元线上,带%标注)。
    marker_series: [(name, vals(美元指数@点 else None), hexcolor, {idx:'0.25%'}), ...]"""
    SIDE = 0.7; LEG_H, DATE_H = 0.70, 0.65
    TOP_BAND = FRAME_MARGIN + LEG_H; BOT_BAND = FRAME_MARGIN + DATE_H + 0.15
    obj_l, obj_t = chart_l + SIDE, chart_t + TOP_BAND
    obj_w, obj_h = chart_w - 2*SIDE, chart_h - TOP_BAND - BOT_BAND
    C = lambda t: f'{{{C_NS}}}{t}'; A = lambda t: f'{{{A_NS}}}{t}'

    cd = ChartData(); cd.categories = dates
    cd.add_series(line_name, line_vals)            # ser0 美元指数(主轴线)
    cd.add_series(area_name, area_vals)            # ser1 国债收益率(→次轴面积)
    for nm, vals, _c, _l in (marker_series or []): # ser2.. 加息/降息(主轴线上的点)
        cd.add_series(nm, vals)
    cs = slide.shapes.add_chart(XL_CHART_TYPE.LINE, Cm(obj_l), Cm(obj_t), Cm(obj_w), Cm(obj_h), cd)
    ch = cs.chart; ch.has_legend = False; ch.has_title = False

    plot = ch._chartSpace.find('.//' + C('plotArea'))
    lineChart = plot.find(C('lineChart'))
    sers = lineChart.findall(C('ser'))
    axIds = lineChart.findall(C('axId'))
    catId1, valId1 = axIds[0].get('val'), axIds[1].get('val')
    catId2, valId2 = '811111111', '822222222'

    def _clear_spPr(ser):
        spPr = ser.find(C('spPr'))
        if spPr is None:
            spPr = ser.makeelement(C('spPr'), {}); ser.find(C('order')).addnext(spPr)
        for c_ in list(spPr): spPr.remove(c_)
        return spPr
    def _marker(ser, hexc, size=9):
        mk = ser.find(C('marker'))
        if mk is None:
            mk = ser.makeelement(C('marker'), {}); ser.find(C('spPr')).addnext(mk)
        for c_ in list(mk): mk.remove(c_)
        etree.SubElement(mk, C('symbol')).set('val', 'circle')
        etree.SubElement(mk, C('size')).set('val', str(size))
        msp = etree.SubElement(mk, C('spPr'))
        sf = etree.SubElement(msp, A('solidFill')); etree.SubElement(sf, A('srgbClr')).set('val', 'FFFFFF')
        ln = etree.SubElement(msp, A('ln')); ln.set('w', '28575')
        sf2 = etree.SubElement(ln, A('solidFill')); etree.SubElement(sf2, A('srgbClr')).set('val', hexc)

    # ser0 美元指数：橙线无点
    sp0 = _clear_spPr(sers[0])
    ln0 = etree.SubElement(sp0, A('ln')); ln0.set('w', str(int(2.5*12700)))
    etree.SubElement(etree.SubElement(ln0, A('solidFill')), A('srgbClr')).set('val', 'F17640')
    mk0 = sers[0].find(C('marker'))
    if mk0 is None:
        mk0 = sers[0].makeelement(C('marker'), {}); sp0.addnext(mk0)
    for c_ in list(mk0): mk0.remove(c_)
    etree.SubElement(mk0, C('symbol')).set('val', 'none')

    # ser1 国债收益率 → areaChart(次轴)
    area_ser = sers[1]; lineChart.remove(area_ser)
    spA = _clear_spPr(area_ser)
    sfA = etree.SubElement(spA, A('solidFill')); scA = etree.SubElement(sfA, A('srgbClr')); scA.set('val', '9DBFE6')
    etree.SubElement(scA, A('alpha')).set('val', '50000')
    areaChart = plot.makeelement(C('areaChart'), {})
    etree.SubElement(areaChart, C('grouping')).set('val', 'standard')
    etree.SubElement(areaChart, C('varyColors')).set('val', '0')
    areaChart.append(area_ser)
    for v in (catId2, valId2): etree.SubElement(areaChart, C('axId')).set('val', v)
    lineChart.addnext(areaChart)

    # ser2.. 加息/降息：留在主轴 lineChart，无线+圆点
    for k, (nm, vals, hexc, _lab) in enumerate(marker_series or []):
        ser = sers[2 + k]
        sp = _clear_spPr(ser); ln = etree.SubElement(sp, A('ln')); etree.SubElement(ln, A('noFill'))
        _marker(ser, hexc, 9)

    # 次轴 valAx2(右,%) + catAx2(隐藏)，国债从0开始
    amax = max(area_vals) if area_vals else 5
    def _ax(tag, axid, cross, delete, pos, fmt=None, mn=None, mx=None):
        ax = plot.makeelement(C(tag), {})
        etree.SubElement(ax, C('axId')).set('val', axid)
        scl = etree.SubElement(ax, C('scaling')); etree.SubElement(scl, C('orientation')).set('val', 'minMax')
        if mx is not None: etree.SubElement(scl, C('max')).set('val', str(mx))
        if mn is not None: etree.SubElement(scl, C('min')).set('val', str(mn))
        etree.SubElement(ax, C('delete')).set('val', delete)
        etree.SubElement(ax, C('axPos')).set('val', pos)
        if fmt:
            nf = etree.SubElement(ax, C('numFmt')); nf.set('formatCode', fmt); nf.set('sourceLinked', '0')
        etree.SubElement(ax, C('majorTickMark')).set('val', 'none')
        etree.SubElement(ax, C('minorTickMark')).set('val', 'none')
        etree.SubElement(ax, C('tickLblPos')).set('val', 'nextTo')
        etree.SubElement(ax, C('crossAx')).set('val', cross)
        if tag == 'catAx':
            etree.SubElement(ax, C('crosses')).set('val', 'max')
        return ax
    import math as _m
    valAx2 = _ax('valAx', valId2, catId2, '0', 'r', '0"%"', 0, int(_m.ceil(amax/2.0)*2))
    catAx2 = _ax('catAx', catId2, valId2, '1', 'b')
    plot.append(valAx2); plot.append(catAx2)

    # 主轴(美元指数)范围
    lo, hi, st = _nice_axis(min(line_vals), max(line_vals))
    ch.value_axis.minimum_scale = lo; ch.value_axis.maximum_scale = hi; ch.value_axis.major_unit = st

    style_chart_xml(ch); _set_plot_area(ch); _set_chart_font(ch)
    # style_chart_xml 会把所有valAx格式改成整数，这里把次轴(国债)恢复成 %
    nf2 = valAx2.find(C('numFmt'))
    if nf2 is not None: nf2.set('formatCode', '0"%"')

    # 黄框
    border = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(chart_l), Cm(chart_t), Cm(chart_w), Cm(chart_h))
    border.fill.background(); border.line.color.rgb = C_YELLOW; border.line.width = Pt(CHART_FRAME_PT); border.shadow.inherit = False
    plot_l = obj_l + PLOT_X*obj_w; plot_t = obj_t + PLOT_Y*obj_h
    plot_w = PLOT_W*obj_w; plot_h = PLOT_H*obj_h; plot_r = plot_l + plot_w

    # 图例
    entries = [(line_name, _pal(0))] + [(nm, RGBColor(int(h[0:2],16),int(h[2:4],16),int(h[4:6],16))) for nm,_,h,_l in (marker_series or [])] + [(area_name, RGBColor(0x9D,0xBF,0xE6))]
    lx, ly = plot_l, chart_t + FRAME_MARGIN
    for nm, col in entries:
        w_name = len(nm)*(15*0.0353)+0.3
        if lx+0.9+w_name > plot_r: lx = plot_l; ly += LEG_H
        sw = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(lx), Cm(ly+LEG_H/2-0.05), Cm(0.7), Cm(0.10))
        sw.fill.solid(); sw.fill.fore_color.rgb = col; sw.line.fill.background(); sw.shadow.inherit=False
        tb = add_tb(slide, nm, lx+0.9, ly, w_name+0.6, LEG_H, size=15, color=C_DARK, align=PP_ALIGN.LEFT)
        try: tb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        except Exception: pass
        lx += 0.9+w_name+0.6

    # 加息/降息 % 标注（按主轴坐标叠加文字框，打在美元指数线上方）
    denom = max(len(dates)-1, 1)
    for nm, vals, hexc, labmap in (marker_series or []):
        for idx, txt in (labmap or {}).items():
            v = line_vals[idx]
            x = plot_l + (idx/denom)*plot_w
            y = plot_t + (1-(v-lo)/((hi-lo) or 1))*plot_h
            add_tb(slide, txt, x-1.0, y-0.95, 2.0, 0.6, size=12, bold=True,
                   color=C_DARK, align=PP_ALIGN.CENTER)

    # 趋势箭头（红升/绿降，叠加）
    up = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Cm(plot_l+plot_w*0.06), Cm(plot_t+plot_h*0.30), Cm(plot_w*0.18), Cm(0.7))
    up.rotation = -38; up.fill.solid(); up.fill.fore_color.rgb = C_RED; up.line.fill.background(); up.shadow.inherit=False
    dn = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Cm(plot_l+plot_w*0.60), Cm(plot_t+plot_h*0.18), Cm(plot_w*0.16), Cm(0.7))
    dn.rotation = 32; dn.fill.solid(); dn.fill.fore_color.rgb = RGBColor(0x2B,0xA8,0x4A); dn.line.fill.background(); dn.shadow.inherit=False

    # 日期 + 数据来源
    date_y = chart_t + chart_h - FRAME_MARGIN - DATE_H
    add_tb(slide, dates[0], plot_l, date_y, 6.5, DATE_H, size=16, bold=True, color=C_DARK, align=PP_ALIGN.LEFT)
    add_tb(slide, dates[-1], plot_r-6.5, date_y, 6.5, DATE_H, size=16, bold=True, color=C_DARK, align=PP_ALIGN.RIGHT)
    add_tb(slide, source, chart_l+chart_w-6.5, chart_t+chart_h+0.05, 6.5, 0.6, size=12, color=C_GRAY, align=PP_ALIGN.RIGHT)
    return cs


# ── 双Y轴 = 两个原生图表叠放（底层面积图Y2右轴 + 顶层透明折线图Y1左轴）──
def _make_transparent(chart):
    """chartSpace + plotArea 改为无填充（顶层图用，露出底层面积）。"""
    cs = chart._chartSpace
    for el in (cs, cs.find(f'.//{{{C_NS}}}plotArea')):
        spPr = el.find(f'{{{C_NS}}}spPr')
        if spPr is None:
            continue
        sf = spPr.find(f'{{{A_NS}}}solidFill')
        if sf is not None:
            spPr.remove(sf)
            idx = list(spPr).index(spPr.find(f'{{{A_NS}}}ln'))if spPr.find(f'{{{A_NS}}}ln') is not None else 0
            nf = spPr.makeelement(f'{{{A_NS}}}noFill', {})
            spPr.insert(0, nf)


def _remove_gridlines(chart):
    plot = chart._chartSpace.find(f'.//{{{C_NS}}}plotArea')
    for gl in plot.findall(f'.//{{{C_NS}}}majorGridlines'):
        gl.getparent().remove(gl)


AREA_BLUE = 'C7DBF0'   # 叠加图底层面积/柱状的浅蓝（低饱和，对齐人工图）

def _set_cross_between(chart, val='midCat'):
    """crossBetween=midCat：点落在刻度上、边对边。叠加图两层都必须设，否则差半格不对齐。"""
    for vax in chart._chartSpace.findall(f'.//{{{C_NS}}}valAx'):
        cb = vax.find(f'{{{C_NS}}}crossBetween')
        if cb is None:
            anchor = vax.find(f'{{{C_NS}}}crosses')
            if anchor is None:
                anchor = vax.find(f'{{{C_NS}}}crossAx')
            cb = vax.makeelement(f'{{{C_NS}}}crossBetween', {})
            anchor.addnext(cb)
        cb.set('val', val)


def add_overlay_dual_axis(slide, chart_l, chart_t, chart_w, chart_h, dates,
                          line_name, line_vals, area_name, area_vals,
                          marker_series=None, bottom_kind='area',
                          area_color=AREA_BLUE, source='数据来源：Wind'):
    """两图叠放实现双Y轴（通用：折线+面积 / 折线+柱状）。
    顶层折线(Y1左轴,透明背景)；底层 bottom_kind='area'面积/'bar'柱状(Y2右轴%)；
    两层同框同绘图区布局+crossBetween=midCat → 上下左右严格对齐。
    marker_series: [(名, vals(线上值/None), hex色, {idx:'0.25%'})] 标记点在顶层折线上。"""
    import math as _m
    SIDE = 0.7; LEG_H, DATE_H = 0.70, 0.65
    TOP_BAND = FRAME_MARGIN + LEG_H; BOT_BAND = FRAME_MARGIN + DATE_H + 0.15
    obj_l, obj_t = chart_l + SIDE, chart_t + TOP_BAND
    obj_w, obj_h = chart_w - 2*SIDE, chart_h - TOP_BAND - BOT_BAND
    C = lambda t: f'{{{C_NS}}}{t}'

    # ── 底层：面积/柱状（Y2，刻度标签放右侧 tickLblPos=high，整数%，从0开始）──
    cd_a = ChartData(); cd_a.categories = dates
    cd_a.add_series(area_name, area_vals)
    btype = XL_CHART_TYPE.COLUMN_CLUSTERED if bottom_kind == 'bar' else XL_CHART_TYPE.AREA
    cs_a = slide.shapes.add_chart(btype, Cm(obj_l), Cm(obj_t), Cm(obj_w), Cm(obj_h), cd_a)
    ch_a = cs_a.chart; ch_a.has_legend = False; ch_a.has_title = False
    ser_a = ch_a.series[0]
    ser_a.format.fill.solid()
    ser_a.format.fill.fore_color.rgb = RGBColor(int(area_color[0:2],16), int(area_color[2:4],16), int(area_color[4:6],16))
    amax = int(_m.ceil(max(area_vals)))
    ch_a.value_axis.minimum_scale = 0
    ch_a.value_axis.maximum_scale = amax
    ch_a.value_axis.major_unit = 1
    style_chart_xml(ch_a)
    va = ch_a._chartSpace.find(f'.//{C_NS and C("valAx")}')
    va.find(C('numFmt')).set('formatCode', '0"%"')
    _ensure(va, C('tickLblPos')).set('val', 'high')     # Y2 刻度放右侧
    _set_cross_between(ch_a, 'midCat')
    _set_plot_area(ch_a); _set_chart_font(ch_a)

    # ── 顶层：折线图（Y1左轴）+ 标记点，整体透明 ──
    cd_l = ChartData(); cd_l.categories = dates
    cd_l.add_series(line_name, line_vals)
    for nm, vals, _c, _l in (marker_series or []):
        cd_l.add_series(nm, vals)
    cs_l = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE_MARKERS if marker_series else XL_CHART_TYPE.LINE,
        Cm(obj_l), Cm(obj_t), Cm(obj_w), Cm(obj_h), cd_l)
    ch_l = cs_l.chart; ch_l.has_legend = False; ch_l.has_title = False
    ch_l.series[0].format.line.color.rgb = _pal(0)
    ch_l.series[0].format.line.width = Pt(CHART_LINE_PT)
    ch_l.series[0].marker.style = XL_MARKER_STYLE.NONE
    for k, (nm, vals, hexc, _l) in enumerate(marker_series or [], start=1):
        ms = ch_l.series[k]
        ms.format.line.fill.background()
        ms.marker.style = XL_MARKER_STYLE.CIRCLE
        ms.marker.size = 9
        ms.marker.format.fill.solid()
        ms.marker.format.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        ms.marker.format.line.color.rgb = RGBColor(int(hexc[0:2],16), int(hexc[2:4],16), int(hexc[4:6],16))
        ms.marker.format.line.width = Pt(2.2)
    lo, hi, st = _nice_axis(min(line_vals), max(line_vals))
    lo, hi, st = int(lo), int(hi), max(int(st), 1)
    ch_l.value_axis.minimum_scale = lo
    ch_l.value_axis.maximum_scale = hi
    ch_l.value_axis.major_unit = st
    style_chart_xml(ch_l)
    _remove_gridlines(ch_l)          # 网格只留底层，避免双重
    _make_transparent(ch_l)          # 透明露出面积
    _set_cross_between(ch_l, 'midCat')   # 与底层同设midCat，两层左右对齐
    _set_plot_area(ch_l); _set_chart_font(ch_l)

    # ── 黄框、图例、%标注、箭头、日期、来源（叠加元素）──
    border = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(chart_l), Cm(chart_t), Cm(chart_w), Cm(chart_h))
    border.fill.background(); border.line.color.rgb = C_YELLOW
    border.line.width = Pt(CHART_FRAME_PT); border.shadow.inherit = False

    plot_l = obj_l + PLOT_X*obj_w; plot_t = obj_t + PLOT_Y*obj_h
    plot_w = PLOT_W*obj_w; plot_h = PLOT_H*obj_h; plot_r = plot_l + plot_w

    entries = [(line_name, _pal(0))]
    entries += [(nm, RGBColor(int(h[0:2],16),int(h[2:4],16),int(h[4:6],16))) for nm,_,h,_l in (marker_series or [])]
    entries += [(area_name, RGBColor(int(area_color[0:2],16), int(area_color[2:4],16), int(area_color[4:6],16)))]
    lx, ly = plot_l, chart_t + FRAME_MARGIN
    for nm, col in entries:
        w_name = len(nm)*(15*0.0353)+0.3
        if lx+0.9+w_name > plot_r: lx = plot_l; ly += LEG_H
        sw = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(lx), Cm(ly+LEG_H/2-0.05), Cm(0.7), Cm(0.10))
        sw.fill.solid(); sw.fill.fore_color.rgb = col; sw.line.fill.background(); sw.shadow.inherit = False
        tb = add_tb(slide, nm, lx+0.9, ly, w_name+0.6, LEG_H, size=15, color=C_DARK, align=PP_ALIGN.LEFT)
        try: tb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        except Exception: pass
        lx += 0.9+w_name+0.6

    denom = max(len(dates)-1, 1)
    for nm, vals, hexc, labmap in (marker_series or []):
        for idx, txt in (labmap or {}).items():
            v = line_vals[idx]
            x = plot_l + (idx/denom)*plot_w
            y = plot_t + (1-(v-lo)/((hi-lo) or 1))*plot_h
            add_tb(slide, txt, x-1.0, y-0.95, 2.0, 0.6, size=12, bold=True,
                   color=C_DARK, align=PP_ALIGN.CENTER)

    up = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Cm(plot_l+plot_w*0.05), Cm(plot_t+plot_h*0.45), Cm(plot_w*0.17), Cm(0.7))
    up.rotation = -35; up.fill.solid(); up.fill.fore_color.rgb = C_RED; up.line.fill.background(); up.shadow.inherit = False
    dn = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Cm(plot_l+plot_w*0.62), Cm(plot_t+plot_h*0.22), Cm(plot_w*0.15), Cm(0.7))
    dn.rotation = 32; dn.fill.solid(); dn.fill.fore_color.rgb = RGBColor(0x2B,0xA8,0x4A); dn.line.fill.background(); dn.shadow.inherit = False

    date_y = chart_t + chart_h - FRAME_MARGIN - DATE_H
    add_tb(slide, dates[0], plot_l, date_y, 6.5, DATE_H, size=16, bold=True, color=C_DARK, align=PP_ALIGN.LEFT)
    add_tb(slide, dates[-1], plot_r-6.5, date_y, 6.5, DATE_H, size=16, bold=True, color=C_DARK, align=PP_ALIGN.RIGHT)
    add_tb(slide, source, chart_l+chart_w-6.5, chart_t+chart_h+0.05, 6.5, 0.6, size=12, color=C_GRAY, align=PP_ALIGN.RIGHT)
    return cs_a, cs_l
