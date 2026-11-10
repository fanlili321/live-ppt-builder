# -*- coding: utf-8 -*-
"""456期：涨指数不涨个股。Wind指数 + 投顾组合净值xlsx + 设计页可编辑重建。"""
import os
import sys
import openpyxl

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE, 'code'))
import ppt_engine as E
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

IMG = os.path.join(BASE, 'examples', 'imgs_456')
COMBO = os.path.join(BASE, 'examples', '投顾组合净值数据_20200101_20260615.xlsx')
OUT = os.path.expanduser('~/Downloads/456期直播PPT_按大纲生成_v2.pptx')
END = '2026-06-12'

# ── 读投顾组合净值 ──
_wb = openpyxl.load_workbook(COMBO, read_only=True, data_only=True)
_ws = _wb.active
def combo(name, col, start=None):
    """col: 2=组合净值 3=基准净值。返回 [{date,close}] 升序。"""
    out = []
    for r in _ws.iter_rows(min_row=2, values_only=True):
        if r[1] == name and r[col] not in (None, ''):
            d = str(r[0])[:10]
            if start and d < start: continue
            try: out.append({'date': d, 'close': float(r[col])})
            except: pass
    out.sort(key=lambda x: x['date']); return out

# ── 读 456 数据需求表 ──
NEED = os.path.join(BASE, 'examples', '【第456期直播】PPT数据需求-快照2.xlsx')
import warnings; warnings.filterwarnings('ignore')
_nwb = openpyxl.load_workbook(NEED, read_only=True, data_only=True)
def need_sheet3():
    """各宽基指数市值 {名称:(平均,中位)}"""
    out={}
    for r in _nwb['3'].iter_rows(min_row=2, values_only=True):
        if r[0] and isinstance(r[1],(int,float)): out[r[0]]=(r[1],r[2])
    return out
def need_listed():
    """Sheet2: [(年, 上市公司总数)]"""
    out=[]
    for r in _nwb['2'].iter_rows(min_row=2, values_only=True):
        if r[0] and isinstance(r[1],(int,float)):
            out.append((str(r[0])[:4], int(r[1])))
    return out
def need_csi_returns():
    """Sheet1: 中证全指成分股 年初至今涨跌幅(%) 列表"""
    out=[]
    for r in _nwb['1'].iter_rows(min_row=6, values_only=True):
        if isinstance(r[3],(int,float)): out.append(r[3]*100)
    return out

# 颜色
ORG=RGBColor(0xF1,0x76,0x40); LORG=RGBColor(0xFF,0xDD,0xC4); BGRAY=RGBColor(0x6B,0x7A,0xA6); LBLUE=RGBColor(0x9D,0xC3,0xE6)
BARBLUE=RGBColor(0xC7,0xDB,0xF0); LINEORG=RGBColor(0xF1,0x76,0x40)

def newslide(prs, body, title):
    sl = prs.slides.add_slide(body)
    for ph in list(sl.placeholders): ph.element.getparent().remove(ph.element)
    E.add_tb(sl, title, E.CONTENT_L, E.CONTENT_TOP, E.CONTENT_W,
             E.text_height_cm(title,E.TITLE_PT,E.CONTENT_W)+0.2,
             size=E.title_fit(title), bold=True, color=E.C_BLACK, align=PP_ALIGN.CENTER)
    return sl

def _btop(title):
    """B: 标题下方统一留白后的正文起始 y。"""
    return E.CONTENT_TOP + E.text_height_cm(title,E.TITLE_PT,E.CONTENT_W)+0.2 + 0.9

def placeholder(prs, body, title, note):
    sl = newslide(prs, body, title)
    E.add_tb(sl, '⚠️ '+note, E.CONTENT_L+2, 8.5, E.CONTENT_W-4, 3,
             size=22, color=E.C_GRAY, align=PP_ALIGN.CENTER)

def build():
    prs = E.Presentation(E.TMPL_PATH)
    body = E.get_layout(prs,1,'自定义版式'); cover = E.get_layout(prs,4,'标题幻灯片')

    # P1 封面
    E.cover_page(prs, cover, '涨指数不涨个股，\n我们该怎么办？')
    # P2/P3/P4 信息图(取第一张)
    E.img_page(prs, body, '股票市场的收益分布不均匀', f'{IMG}/v01.png')
    E.img_page(prs, body, '美股的长期数据：少数股票贡献大部分回报', f'{IMG}/v02.png')
    E.img_page(prs, body, 'A股也有类似的特点', f'{IMG}/v04.png')
    # P5 中证全指成分股涨幅分布(条形图)
    _p5(prs, body)
    # P6 A股上市公司数量年度(柱状图)
    _p6(prs, body)
    # P7 两种投资方法：可编辑对比框+气泡
    E.box_page(prs, body, '两种投资方法', [
        ('主动投资','核心：挑选长期盈利能力更强的股票\n代表：巴菲特、彼得林奇\n难点：多数普通投资者不具备挑选能力'),
        ('被动投资','核心：把整个市场买下来（如宽基指数基金）\n代表：宽基指数基金等\n难点：需低估买入、控制成本'),
    ], '「涨指数不涨个股」其实最好解决——买宽基指数基金就好了。这也是欧美日韩及A股指数化投资成为主力的原因。')
    # P8 四类指数：四个圆(素材库色，两侧留边、圆间距收紧)
    _p8(prs, body)
    # P9 什么是宽基：树状(可编辑，照原版重做、不遮挡)
    _p9(prs, body)
    # P10 代表性宽基指数表格(可编辑, Sheet3 市值, 截至2026-06-12)
    mv=need_sheet3()  # {指数:(平均,中位)}
    meta={'沪深300':('大盘股','规模最大1-300'),'中证500':('中盘股','301-800'),
          '中证800':('大中盘股','1-800'),'中证1000':('小盘股','801-1800'),
          '中证2000':('小微盘股','1801-3800'),'中证全指':('市场整体','沪深300+500+1000+2000')}
    p10_headers=['指数','含义','成分股','市值平均(亿)','市值中位数(亿)']
    p10_rows=[[k,meta[k][0],meta[k][1],f'{mv[k][0]:,.2f}',f'{mv[k][1]:,.2f}'] for k in meta if k in mv]
    _table_page(prs, body, 'A股市场上比较有代表性的宽基指数', p10_headers, p10_rows,
                widths=[4.5,4.5,9.0,7.0,7.0], src='数据来源：Wind，截至2026-06-12')
    # P11 中证800 双线(Wind,周)
    _wind_dual(prs, body, '投资宽基指数，选哪个更合适？','（1）中证800（=沪深300+中证500）',
               '000906.CSI','H00906.CSI','中证800','中证800全收益','2007-01-15',END,'W')
    # P12 指数增强 双线(组合xlsx日 + 中证800 Wind日)
    ie = combo('螺丝钉金钉宝指数增强C',2,'2024-09-02')
    _p12(prs, body, ie)
    # P13 左右两图
    fm = combo('螺丝钉金钉宝全市场指数C1',2,'2024-02-28'); fmb = combo('螺丝钉金钉宝全市场指数C1',3,'2024-02-28')
    _p13(prs, body, fm, fmb)
    # P14 微笑曲线(可编辑重建)
    _p14(prs, body)
    # P15 左:星级含义表(可编辑) 右:估值表截图
    _p15(prs, body)
    # P16 总结
    E.arrows_page(prs, body, '总结', [
        '股票市场的收益分布不均匀，导致"涨指数不涨个股"真实存在。',
        '个股投资越来越难，普通投资者更适合通过指数化投资参与市场。',
        '可以考虑从中证800等宽基指数入手，在低估时投资、持有到高估止盈。',
        '可参考螺丝钉星级及指数估值表，快速判断市场整体及单个指数的估值。'], pt=26, spread=True)

    for _ in range(8): E.remove_slide(prs,0)
    prs.save(OUT); print(f'\n✅ 完成！{len(prs.slides)} 张 → {OUT}')

# ── 辅助：通用表格页 ──
def _table_page(prs, body, title, headers, rows, widths, src='', topgap=None, ret=False):
    sl = newslide(prs, body, title)
    nr=len(rows)+1; nc=len(headers)
    # B/I: 表格起点压在标题下方留白处，绝不顶标题
    t_h=E.text_height_cm(title,E.TITLE_PT,E.CONTENT_W)+0.2
    top=E.CONTENT_TOP+t_h+1.0
    # J: 纵向铺满到页脚上方(留出来源行)
    tbl_h=E.CONTENT_BOTTOM-top-0.9
    gt=sl.shapes.add_table(nr,nc,Cm(E.CONTENT_L),Cm(top),Cm(E.CONTENT_W),Cm(tbl_h)).table
    tot=sum(widths)
    for c,w in enumerate(widths): gt.columns[c].width=Cm(w/tot*E.CONTENT_W)
    for r in range(nr): gt.rows[r].height=Cm(tbl_h/nr)
    def cell(r,c,t,bold=False,color=E.C_DARK,fill=None,al=PP_ALIGN.CENTER):
        cl=gt.cell(r,c); cl.fill.solid(); cl.fill.fore_color.rgb=fill or RGBColor(0xFF,0xFF,0xFF)
        cl.margin_top=Cm(0.05); cl.margin_bottom=Cm(0.05); cl.vertical_anchor=MSO_ANCHOR.MIDDLE
        tf=cl.text_frame; tf.word_wrap=True; p=tf.paragraphs[0]; p.alignment=al   # K: 长内容换行
        run=p.add_run(); run.text=t; run.font.size=Pt(18); run.font.bold=bold; run.font.color.rgb=color; E.set_run_font(run)
    for c,h in enumerate(headers): cell(0,c,h,True,RGBColor(0xFF,0xFF,0xFF),E.C_MAIN)
    for r,row in enumerate(rows,1):
        bg=RGBColor(0xF5,0xF5,0xF5) if r%2==0 else RGBColor(0xFF,0xFF,0xFF)
        for c,v in enumerate(row): cell(r,c,v,bold=(c==0),fill=bg,al=(PP_ALIGN.LEFT if c==1 else PP_ALIGN.CENTER))
    if src: E.add_tb(sl,src,E.CONTENT_L+E.CONTENT_W-12,top+tbl_h+0.15,12,0.5,size=12,color=E.C_GRAY,align=PP_ALIGN.RIGHT)
    return sl if ret else None

# ── 容错取数：按年分段，每年多轮重试，跳过死段 ──
def _kl(code, start, end, freq):
    import datetime as _dt
    y0=int(start[:4]); y1=int(end[:4]); out=[]
    for y in range(y0, y1+1):
        s=f'{y}-01-01' if y>y0 else start
        e=f'{y}-12-31' if y<y1 else end
        got=None
        for _ in range(8):
            try: got=E.fetch_kline(code,s,e,freq); break
            except Exception: continue
        if got: out.extend(got)
    out.sort(key=lambda x:x['date']); return out

# ── 辅助：Wind双线页(带副标题) ──
def _wind_dual(prs, body, title, subtitle, c1, c2, n1, n2, start, end, freq):
    d1=_kl(c1,start,end,freq); d2=_kl(c2,start,end,freq)
    def pct(d): b=d[0]['close']; return [{'date':x['date'],'close':(x['close']/b-1)*100} for x in d]
    sl_list=[{'name':n1,'dates':[x['date'] for x in d1],'values':[v['close'] for v in pct(d1)]},
             {'name':n2,'dates':[x['date'] for x in d2],'values':[v['close'] for v in pct(d2)]}]
    E.multi_chart_page(prs, body, title, sl_list, points=[subtitle], y_pct=True)

# ── 辅助：组合(内联)+Wind指数 多线 ──
def _multi_inline_wind(prs, body, title, subtitle, inline_series, wind_series, start, end, freq):
    sl_list=[]
    for nm,d in inline_series:
        d=[x for x in d if start<=x['date']<=end]; b=d[0]['close']
        sl_list.append({'name':nm,'dates':[x['date'] for x in d],'values':[(x['close']/b-1)*100 for x in d]})
    for nm,c in wind_series:
        d=E.fetch_kline(c,start,end,freq); b=d[0]['close']
        sl_list.append({'name':nm,'dates':[x['date'] for x in d],'values':[(x['close']/b-1)*100 for x in d]})
    E.multi_chart_page(prs, body, title, sl_list, points=[subtitle] if subtitle else None, y_pct=True)

# ── P13 左右两图 ──
def _p13(prs, body, fm, fmb):
    sl = newslide(prs, body, '投资宽基指数，选哪个更合适？')
    E.add_arrow_points(sl, ['（2）中证全指（=沪深300+中证500+中证1000+中证2000）；2025年来首现中证全指基金。'],
                       E.CONTENT_TOP+1.5, left=E.CONTENT_L, width=E.CONTENT_W)
    ct=E.CONTENT_TOP+3.3; chh=E.CONTENT_BOTTOM-0.6-ct; halfw=(E.CONTENT_W-0.8)/2
    # 左：中证全指+全收益(Wind,周)
    d1=_kl('000985.CSI','2004-12-31',END,'W'); d2=_kl('H00985.CSI','2004-12-31',END,'W')
    def pct(d): b=d[0]['close']; return [(x['close']/b-1)*100 for x in d]
    left=[{'name':'中证全指','dates':[x['date'] for x in d1],'values':pct(d1)},
          {'name':'中证全指全收益','dates':[x['date'] for x in d2],'values':pct(d2)}]
    E.add_multi_line_chart(sl, E.CONTENT_L, ct, halfw, chh, left, y_pct=True)
    # 右：金钉宝全市场组合+业绩基准(内联,日)
    fb=fm[0]['close']; bb=fmb[0]['close']
    right=[{'name':'金钉宝全市场指数组合','dates':[x['date'] for x in fm],'values':[(x['close']/fb-1)*100 for x in fm]},
           {'name':'业绩基准','dates':[x['date'] for x in fmb],'values':[(x['close']/bb-1)*100 for x in fmb]}]
    E.add_multi_line_chart(sl, E.CONTENT_L+halfw+0.8, ct, halfw, chh, right, y_pct=True)
    E.add_tb(sl,'数据来源：Wind、螺丝钉投顾组合净值，截至2026-06-12',
             E.CONTENT_L+E.CONTENT_W-14, E.CONTENT_BOTTOM-0.45, 14, 0.45,
             size=11, color=E.C_GRAY, align=PP_ALIGN.RIGHT)

# ── P14 微笑曲线：可编辑色带 + 曲线 + 圆点 + 标签 ──
def _p14(prs, body):
    from pptx.enum.shapes import MSO_CONNECTOR
    sl = newslide(prs, body, '投资方法：低估买入，正常估值持有，高估卖出')
    l, t, w, h = E.CONTENT_L+1.8, E.CONTENT_TOP+2.0, E.CONTENT_W-5.2, 10.8
    band_h = h / 3
    bands = [
        (t, RGBColor(0xF9,0xDC,0xB8), '高估卖出'),
        (t+band_h, RGBColor(0xFB,0xEF,0xC5), '正常估值\n持有'),
        (t+2*band_h, RGBColor(0xD8,0xF0,0xCB), '低估买入'),
    ]
    for y, col, label in bands:
        shp = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(l), Cm(y), Cm(w), Cm(band_h))
        shp.fill.solid(); shp.fill.fore_color.rgb = col
        shp.line.color.rgb = RGBColor(0xE2,0xE2,0xE2); shp.line.width = Pt(0.6)
        shp.shadow.inherit = False
    # 轻网格
    for i in range(1, 9):
        x = l + w*i/9
        ln = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(x), Cm(t), Cm(x), Cm(t+h))
        ln.line.color.rgb = RGBColor(0xEF,0xEF,0xEF); ln.line.width = Pt(0.6)
    for i in range(1, 6):
        y = t + h*i/6
        ln = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(l), Cm(y), Cm(l+w), Cm(y))
        ln.line.color.rgb = RGBColor(0xEF,0xEF,0xEF); ln.line.width = Pt(0.6)

    xs = [0.03,0.12,0.22,0.32,0.42,0.52,0.62,0.72,0.82,0.93]
    ys = [0.03,0.27,0.50,0.65,0.75,0.79,0.75,0.62,0.35,0.03]
    pts = [(l + w*x, t + h*y) for x,y in zip(xs,ys)]
    for (x1,y1),(x2,y2) in zip(pts, pts[1:]):
        ln = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Cm(x1), Cm(y1), Cm(x2), Cm(y2))
        ln.line.color.rgb = RGBColor(0x3A,0x3A,0x3A); ln.line.width = Pt(2.2)
    for x,y in pts[1:-1]:
        dot = sl.shapes.add_shape(MSO_SHAPE.OVAL, Cm(x-0.13), Cm(y-0.13), Cm(0.26), Cm(0.26))
        dot.fill.solid(); dot.fill.fore_color.rgb = RGBColor(0xFF,0xFF,0xFF)
        dot.line.color.rgb = E.C_MAIN; dot.line.width = Pt(1.4)
        dot.shadow.inherit = False

    label_x = l + w + 0.45
    for y, _, label in bands:
        tag = sl.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Cm(label_x), Cm(y+band_h/2-0.55), Cm(3.2), Cm(1.1))
        tag.fill.solid(); tag.fill.fore_color.rgb = RGBColor(0xF7,0xD9,0xB4)
        tag.line.color.rgb = RGBColor(0xED,0xBD,0x83); tag.line.width = Pt(1.0)
        tag.shadow.inherit = False
        tb = E.add_tb(sl, label, label_x+0.15, y+band_h/2-0.45, 2.5, 0.9,
                      size=16, bold=True, color=E.C_DARK, align=PP_ALIGN.CENTER)
        tb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    E.add_tb(sl,'数据来源：银行螺丝钉直播大纲/历史页重绘',
             E.CONTENT_L+E.CONTENT_W-10, E.CONTENT_BOTTOM-0.45, 10, 0.45,
             size=11, color=E.C_GRAY, align=PP_ALIGN.RIGHT)

# ── P12 指数增强组合 + 中证800，叠加申购开/关标注 ──
def _p12(prs, body, ie):
    ie=[x for x in ie if '2024-09-02'<=x['date']<=END]
    b=ie[0]['close']
    s_ie={'name':'指数增强组合','dates':[x['date'] for x in ie],'values':[(x['close']/b-1)*100 for x in ie]}
    d=_kl('000906.CSI','2024-09-02',END,'D'); bb=d[0]['close']
    s_idx={'name':'中证800','dates':[x['date'] for x in d],'values':[(x['close']/bb-1)*100 for x in d]}
    series=[s_ie,s_idx]
    sl=newslide(prs, body, '指数增强基金的表现')
    cl,ct=E.CONTENT_L, E.CONTENT_TOP+1.6; cw=E.CONTENT_W; chh=E.CONTENT_BOTTOM-0.6-ct
    E.add_multi_line_chart(sl, cl, ct, cw, chh, series, y_pct=True)
    SIDE=0.7; LEG_H=0.70; DATE_H=0.65; FM=E.FRAME_MARGIN
    TOP_BAND=FM+LEG_H; BOT_BAND=FM+DATE_H+0.15
    obj_l=cl+SIDE; obj_t=ct+TOP_BAND; obj_w=cw-2*SIDE; obj_h=chh-TOP_BAND-BOT_BAND
    plot_l=obj_l+E.PLOT_X*obj_w; plot_w=E.PLOT_W*obj_w
    plot_t=obj_t+E.PLOT_Y*obj_h; plot_h=E.PLOT_H*obj_h
    common=set(series[0]['dates'])
    for s in series[1:]: common&=set(s['dates'])
    cats=[x for x in series[0]['dates'] if x in common]; N=len(cats)
    events=[('2026-01-12','权益组合\n关闭申购',0),('2026-03-23','打开申购',1),('2026-04-14','关闭申购',0)]
    from pptx.dml.color import RGBColor; from pptx.enum.shapes import MSO_SHAPE
    GREY=RGBColor(0x88,0x88,0x88); GREEN=RGBColor(0x2B,0xA8,0x4A)
    for ed,label,lvl in events:
        idx=next((i for i,c in enumerate(cats) if c>=ed), None)
        if idx is None: continue
        x=plot_l+(idx+0.5)/N*plot_w
        ln=sl.shapes.add_shape(MSO_SHAPE.RECTANGLE,Cm(x-0.015),Cm(plot_t),Cm(0.03),Cm(plot_h))
        ln.fill.solid(); ln.fill.fore_color.rgb=GREY; ln.line.fill.background(); ln.shadow.inherit=False
        ly=plot_t+0.1+lvl*1.3
        E.add_tb(sl,label,x-1.6,ly,3.2,1.2,size=12,bold=True,
                 color=(E.C_RED if '关闭' in label else GREEN),align=PP_ALIGN.CENTER)
    E.add_tb(sl,'数据来源：螺丝钉投顾组合净值、Wind',cl+cw-12,ct+chh-0.5,12,0.5,size=12,color=E.C_GRAY,align=PP_ALIGN.RIGHT)

# ── 原生柱状图 ──
def _bar_chart(sl, cats, vals, l, t, w, h, colors, src, lbl_size=12, cat_size=12, show_labels=True):
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.chart.data import ChartData
    SIDE=0.7; FM=E.FRAME_MARGIN; DATE_H=0.0
    obj_l=l+SIDE; obj_t=t+FM+0.1; obj_w=w-2*SIDE; obj_h=h-2*FM-0.2
    cd=ChartData(); cd.categories=cats; cd.add_series('数量', vals)
    cs=sl.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,Cm(obj_l),Cm(obj_t),Cm(obj_w),Cm(obj_h),cd)
    ch=cs.chart; ch.has_legend=False; ch.has_title=False
    plot=ch.plots[0]; plot.gap_width=60; ser=plot.series[0]
    for i,pt in enumerate(ser.points):
        pt.format.fill.solid(); pt.format.fill.fore_color.rgb=colors[i % len(colors)]
        pt.format.line.fill.background()
    ymin,ymax,ystep=E._nice_axis(0, max(vals))
    ch.value_axis.minimum_scale=0; ch.value_axis.maximum_scale=ymax; ch.value_axis.major_unit=ystep
    ch.value_axis.tick_labels.font.size=Pt(cat_size)
    ch.value_axis.has_major_gridlines=True
    ch.category_axis.tick_labels.font.size=Pt(cat_size)
    plot.has_data_labels=show_labels   # E: 标签会重叠时(柱多)关闭
    if show_labels:
        dl=plot.data_labels; dl.font.size=Pt(lbl_size); dl.font.bold=True; dl.number_format='#,##0'; dl.number_format_is_linked=False
    E._set_chart_font(ch)
    bd=sl.shapes.add_shape(MSO_SHAPE.RECTANGLE,Cm(l),Cm(t),Cm(w),Cm(h))
    bd.fill.background(); bd.line.color.rgb=E.C_YELLOW; bd.line.width=Pt(E.CHART_FRAME_PT); bd.shadow.inherit=False
    E.add_tb(sl,src,l+w-12,t+h+0.05,12,0.5,size=12,color=E.C_GRAY,align=PP_ALIGN.RIGHT)
    return cs

# ── P5 中证全指成分股涨幅分布 ──
def _p5(prs, body):
    import statistics
    vals=need_csi_returns(); n=len(vals); idx=7.79
    edges=[(-1e9,-30),(-30,-10),(-10,0),(0,10),(10,30),(30,50),(50,100),(100,1e9)]
    labels=['跌超\n30%','跌10~\n30%','跌0~\n10%','涨0~\n10%','涨10~\n30%','涨30~\n50%','涨50~\n100%','涨超\n100%']
    counts=[sum(1 for x in vals if lo<=x<hi) for lo,hi in edges]
    down=sum(1 for x in vals if x<0); win=sum(1 for x in vals if x>idx)
    med=statistics.median(vals)
    cols=[BARBLUE,BARBLUE,BARBLUE,LINEORG,LINEORG,LINEORG,LINEORG,LINEORG]
    sl=newslide(prs, body, '涨指数不涨个股')
    txt=(f'2026年以来中证全指上涨 +{idx:.2f}%，但 {n:,} 只成分股中 {down:,} 只（{down/n*100:.0f}%）下跌、'
         f'涨跌幅中位数 −{abs(med):.1f}%，仅 {win:,} 只（{win/n*100:.0f}%）跑赢指数——典型的"涨指数不涨个股"。')
    y=E.add_arrow_points(sl, [txt], _btop('涨指数不涨个股'), left=E.CONTENT_L, width=E.CONTENT_W)
    ct=y+0.6; chh=E.CONTENT_BOTTOM-0.6-ct
    _bar_chart(sl, labels, counts, E.CONTENT_L, ct, E.CONTENT_W, chh, cols,
               '数据来源：Wind，中证全指成分股年初至今涨跌幅', lbl_size=13, cat_size=13)

# ── P6 A股上市公司数量年度 ──
def _p6(prs, body):
    data=need_listed()
    cats=[y for y,_ in data]; vals=[v for _,v in data]
    tt='A股上市公司数量变多，结构性牛市更常见'
    sl=newslide(prs, body, tt)
    y=E.add_arrow_points(sl, ['上市公司越来越多，单只个股越来越难选；指数则能持续纳入优质新公司，更适合长期投资。'],
                       _btop(tt), left=E.CONTENT_L, width=E.CONTENT_W)
    ct=y+0.6; chh=E.CONTENT_BOTTOM-0.6-ct
    _bar_chart(sl, cats, vals, E.CONTENT_L, ct, E.CONTENT_W, chh, [ORG],
               '数据来源：Wind，截至2026-06-12', cat_size=11, show_labels=False)

# ── P8 四类指数(素材库色，两侧留边、圆间距收紧、不贴边) ──
def _p8(prs, body):
    sl=newslide(prs, body, 'A股常见的四类指数')
    # 素材库：宽基=橙#F17640白字, 策略=浅橙#FCE6D2深字, 行业=灰蓝#6B7AA6白字, 主题=浅蓝#9DC3E6白字
    items=[('宽基指数',ORG,RGBColor(0xFF,0xFF,0xFF)),('策略指数',RGBColor(0xFC,0xE6,0xD2),E.C_DARK),
           ('行业指数',BGRAY,RGBColor(0xFF,0xFF,0xFF)),('主题指数',LBLUE,RGBColor(0xFF,0xFF,0xFF))]
    side=1.5; gap=1.0; n=4
    d=(E.CONTENT_W-2*side-(n-1)*gap)/n            # ≈6.6，圆更大、间距收紧
    cy=8.3
    for i,(t,c,tc) in enumerate(items):
        x=E.CONTENT_L+side+i*(d+gap)
        o=sl.shapes.add_shape(MSO_SHAPE.OVAL,Cm(x),Cm(cy),Cm(d),Cm(d))
        o.fill.solid(); o.fill.fore_color.rgb=c; o.line.fill.background(); o.shadow.inherit=False
        tb=E.add_tb(sl,t,x,cy,d,d,size=24,bold=True,color=tc,align=PP_ALIGN.CENTER)
        tb.text_frame.vertical_anchor=MSO_ANCHOR.MIDDLE

# ── P9 什么是宽基：照原版大纲1:1(金橙根圆+浅橙无边叶子+橙色扇形连线) ──
def _p9(prs, body):
    tt='什么是宽基指数？'
    sl=prs.slides.add_slide(body)
    for ph in list(sl.placeholders): ph.element.getparent().remove(ph.element)
    # 标题左对齐(原版样式)，避免居中标题压到右侧树
    E.add_tb(sl, tt, E.CONTENT_L, E.CONTENT_TOP, 13.0,
             E.text_height_cm(tt,E.TITLE_PT,13.0)+0.2, size=E.TITLE_PT, bold=True, color=E.C_BLACK, align=PP_ALIGN.LEFT)
    y0=_btop(tt)
    E.add_tb(sl,'通常包含各个行业的股票，覆盖范围很"宽"。\n可能在聚焦的股票规模上会有区分，如：',
             E.CONTENT_L,y0,12.0,2.0,size=20,color=E.C_DARK,align=PP_ALIGN.LEFT)
    rows=[('沪深300','大盘股'),('中证500','中盘股'),('中证800','大中盘股'),
          ('中证1000','小盘股'),('中证2000','小微盘股'),('中证全指','市场整体')]
    GOLD=RGBColor(0xF2,0xAE,0x52); PALE=RGBColor(0xFC,0xE6,0xD2)
    # 叶子：右侧竖排
    top=E.CONTENT_TOP+1.4; bot=E.CONTENT_BOTTOM-0.2; n=len(rows)
    slot=(bot-top)/n; bh=slot-0.5
    lx=15.5; lw=E.CONTENT_L+E.CONTENT_W-lx
    leaf_cys=[]
    for i,(a,b) in enumerate(rows):
        yy=top+i*slot
        bar=sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Cm(lx),Cm(yy),Cm(lw),Cm(bh))
        bar.fill.solid(); bar.fill.fore_color.rgb=PALE; bar.line.fill.background(); bar.shadow.inherit=False
        tb=E.add_tb(sl,'',lx+0.7,yy,lw-1.2,bh,size=18,bold=True,color=E.C_DARK,align=PP_ALIGN.LEFT)
        tb.text_frame.vertical_anchor=MSO_ANCHOR.MIDDLE
        p=tb.text_frame.paragraphs[0]
        r1=p.add_run(); r1.text=f'{a}：'; r1.font.size=Pt(18); r1.font.bold=True; r1.font.color.rgb=ORG; E.set_run_font(r1)
        r2=p.add_run(); r2.text=b; r2.font.size=Pt(18); r2.font.bold=True; r2.font.color.rgb=E.C_DARK; E.set_run_font(r2)
        leaf_cys.append(yy+bh/2)
    # 根节点：金橙圆、深色字、竖直居中
    nd=4.6; ncx=E.CONTENT_L+1.5; ncy=top+(bot-top)/2-nd/2; rootcy=ncy+nd/2
    node=sl.shapes.add_shape(MSO_SHAPE.OVAL,Cm(ncx),Cm(ncy),Cm(nd),Cm(nd))
    node.fill.solid(); node.fill.fore_color.rgb=GOLD; node.line.fill.background(); node.shadow.inherit=False
    tb=E.add_tb(sl,'宽基指数',ncx,ncy,nd,nd,size=22,bold=True,color=E.C_DARK,align=PP_ALIGN.CENTER); tb.text_frame.vertical_anchor=MSO_ANCHOR.MIDDLE
    # 橙色扇形连线：根右缘 → 每个叶子左缘(肘形)
    busx=ncx+nd+1.2
    def seg(x,y,w,h):
        ln=sl.shapes.add_shape(MSO_SHAPE.RECTANGLE,Cm(x),Cm(y),Cm(max(w,0.04)),Cm(max(h,0.04)))
        ln.fill.solid(); ln.fill.fore_color.rgb=ORG; ln.line.fill.background(); ln.shadow.inherit=False
    seg(ncx+nd, rootcy-0.02, busx-(ncx+nd), 0.04)              # 根→总线
    seg(busx-0.02, min(leaf_cys+[rootcy]), 0.04, max(leaf_cys+[rootcy])-min(leaf_cys+[rootcy]))  # 竖总线
    for cyl in leaf_cys: seg(busx, cyl-0.02, lx-busx, 0.04)     # 总线→各叶子

# ── P15 标题+要点+星级含义表(左) ‖ 估值表+三色块标签(右)，照大纲版式 ──
def _p15(prs, body):
    WHITE=RGBColor(0xFF,0xFF,0xFF)
    sl=prs.slides.add_slide(body)
    for ph in list(sl.placeholders): ph.element.getparent().remove(ph.element)
    # 标题(左对齐,两行)
    E.add_tb(sl,'螺丝钉星级&指数估值表：\n判断市场整体&单个指数的估值',
             0.8,0.7,17.0,2.8,size=30,bold=True,color=E.C_BLACK,align=PP_ALIGN.LEFT)
    # 两个要点(橙色箭头)
    E.add_arrow_points(sl,['螺丝钉星级：判断市场整体是便宜还是贵。',
                           '指数估值表：判断单个指数是便宜还是贵。'],3.9,left=0.9,width=16.5)
    # 右侧估值表图(566x1466)
    img=f'{IMG}/valuation_20260612.png'
    ih=15.0; iw=ih*566/1466; ix=21.6; iy=1.5
    sl.shapes.add_picture(img, Cm(ix), Cm(iy), height=Cm(ih))
    E.add_tb(sl,'截至2026-06-12，当前3.5星',ix-0.8,iy+ih+0.02,iw+1.6,0.5,size=10,color=E.C_GRAY,align=PP_ALIGN.CENTER)
    # 最右三色块标签(对齐估值表 绿/橙/红 三段)
    lx=ix+iw+0.45; lw=E.CONTENT_L+E.CONTENT_W-lx
    bands=[('低估\n适合定投',RGBColor(0x5B,0x9B,0xD5),iy+0.13*ih),
           ('正常估值\n暂停定投、继续持有',RGBColor(0xF1,0x76,0x40),iy+0.52*ih),
           ('高估\n分批止盈',RGBColor(0xC2,0x47,0x85),iy+0.90*ih)]
    for txt,clr,cy in bands:
        bh=2.4
        box=sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Cm(lx),Cm(cy-bh/2),Cm(lw),Cm(bh))
        box.fill.solid(); box.fill.fore_color.rgb=clr; box.line.fill.background(); box.shadow.inherit=False
        tb=E.add_tb(sl,txt,lx,cy-bh/2,lw,bh,size=13,bold=True,color=WHITE,align=PP_ALIGN.CENTER)
        tb.text_frame.vertical_anchor=MSO_ANCHOR.MIDDLE
    # 左侧星级含义表
    rows=[['5星级-5.9星级','股票基金投资最好的阶段'],['4星级-4.9星级','熊市通常会达到，会有一部分低估品种'],
          ['3星级-3.9星级','低估品种很少，大部分品种正常和高估'],['2星级-2.9星级','牛市中后期，大部分品种进入高估'],
          ['1星级-1.9星级','牛市巅峰泡沫阶段']]
    tl,tt,tw=0.9, 7.6, 19.8
    nr=len(rows)+1; th=1.45*nr
    gt=sl.shapes.add_table(nr,2,Cm(tl),Cm(tt),Cm(tw),Cm(th)).table
    gt.columns[0].width=Cm(tw*0.34); gt.columns[1].width=Cm(tw*0.66)
    def cell(r,c,t,bold,color,fill,al):
        cl=gt.cell(r,c); cl.fill.solid(); cl.fill.fore_color.rgb=fill; cl.vertical_anchor=MSO_ANCHOR.MIDDLE
        cl.margin_top=Cm(0.04); cl.margin_bottom=Cm(0.04)
        p=cl.text_frame.paragraphs[0]; p.alignment=al; cl.text_frame.word_wrap=True
        run=p.add_run(); run.text=t; run.font.size=Pt(15); run.font.bold=bold; run.font.color.rgb=color; E.set_run_font(run)
    cell(0,0,'螺丝钉星级',True,WHITE,E.C_MAIN,PP_ALIGN.CENTER)
    cell(0,1,'含义',True,WHITE,E.C_MAIN,PP_ALIGN.CENTER)
    for r,(a,b) in enumerate(rows,1):
        bg=RGBColor(0xF5,0xF5,0xF5) if r%2==0 else WHITE
        cell(r,0,a,True,E.C_DARK,bg,PP_ALIGN.CENTER); cell(r,1,b,False,E.C_DARK,bg,PP_ALIGN.LEFT)

build()
