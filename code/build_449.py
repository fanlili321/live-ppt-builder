# -*- coding: utf-8 -*-
"""449期完整版：xlsx真实数据为主，2016-2017用Wind，星级/信号板/止盈/估值表/白酒占位。"""
import os
import sys
import openpyxl
from datetime import datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ppt_engine as E

XLSX = os.path.expanduser('~/Downloads/【第449期直播】数据需求20260518-快照4.xlsx')
wb = openpyxl.load_workbook(XLSX, data_only=True, read_only=True)

def series(sheet, dc, vc, sr):
    out = []
    for r in wb[sheet].iter_rows(min_row=sr, values_only=True):
        d, v = r[dc], r[vc]
        if d is None or v in (None, '', 0): continue
        try: out.append({'date': str(d)[:10], 'close': float(v)})
        except (TypeError, ValueError): continue
    return out

def pct(s):
    b = s[0]['close']; return [{'date': x['date'], 'close': (x['close']/b-1)*100} for x in s]
def msl(pairs):
    return [{'name': n, 'dates': [x['date'] for x in s], 'values': [y['close'] for y in pct(s)]} for n, s in pairs]
def days(a, b):
    return (datetime.strptime(b[:10], '%Y-%m-%d') - datetime.strptime(a[:10], '%Y-%m-%d')).days

# ── 数据 ──
p2  = [('全球股票市场', series('1',0,1,3)), ('美股全市场', series('1',0,2,3)),
       ('中证全指全收益', series('1',0,3,3)), ('港股恒生全收益', series('1',0,4,3))]
p5  = [('中证全指', series('3',0,1,5)), ('中证全指全收益', series('3',0,2,5))]
p5_table = {'headers':['归母净利润(亿元)','2020','2021','2022','2023','2024','2025','2026Q1'],
            'rows':[['盈利','39468.66','48103.93','52385.85','52598.97','52478.44','53692.17','15915.64'],
                    ['同比','4.24%','21.88%','8.90%','0.41%','-0.23%','2.34%','7.02%']]}
p6_line = series('3',0,1,5)
p9  = [('中证1000全收益', series('5',0,1,5)), ('沪深300全收益', series('5',0,2,5)), ('中证全指全收益', series('5',0,3,5))]
p12 = [('科创50全收益', series('6',0,1,5)), ('中证全指全收益', series('6',0,2,5))]
p15 = [('中证全指证券公司全收益', series('7',0,1,5))]
p16 = [('中证1000全收益', series('8',0,1,5)), ('创业板指全收益', series('8',0,2,5)), ('中证红利全收益', series('8',0,3,5)),
       ('300价值全收益', series('8',0,4,5)), ('800消费全收益', series('8',0,5,5))]
p18 = [('365天投顾组合', series('9',0,1,2)), ('偏债混合基金指数', series('9',0,2,2)), ('二级债基指数', series('9',0,3,2))]

# P4 双轴
p4_usd = series('2',0,1,5); p4_yld = series('2',4,5,5)
FOMC = [('2022-03-16','hike','0.25%'),('2022-05-04','hike','0.50%'),('2022-06-15','hike','0.75%'),
        ('2022-07-27','hike','0.75%'),('2022-09-21','hike','0.75%'),('2022-11-02','hike','0.75%'),
        ('2022-12-14','hike','0.50%'),('2023-02-01','hike','0.25%'),('2023-05-03','hike','0.25%'),
        ('2023-07-26','hike','0.25%'),('2024-09-18','cut','-0.50%'),('2024-11-07','cut','-0.25%'),
        ('2024-12-18','cut','-0.25%'),('2025-09-17','cut','-0.25%')]
def _snap(d, ser):
    ds={x['date'] for x in ser}
    if d in ds: return d
    td=datetime.strptime(d,'%Y-%m-%d'); return min(ser,key=lambda x:abs((datetime.strptime(x['date'],'%Y-%m-%d')-td).days))['date']
p4_marks=[{'date':_snap(d,p4_usd),'type':tp,'bp':bp} for d,tp,bp in FOMC]

# P8 6轮牛熊（Sheet4 权威）
ws4 = wb['4']; p8_rows = []
for r in ws4.iter_rows(min_row=3, max_row=8, values_only=True):
    rn = r[0]
    if not rn: continue
    bull_s, bull_e, bull_chg = str(r[8])[:10], str(r[9])[:10], r[10]
    if bull_s and bull_s != 'None':
        p8_rows.append([rn, '牛', f'{bull_s} ~ {bull_e}', str(days(bull_s, bull_e)), f'+{float(bull_chg)*100:.2f}%'])
    bear_s = str(r[1])[:10]
    if bear_s and bear_s not in ('None',''):
        p8_rows.append([rn, '熊', f'{bear_s} ~ {str(r[2])[:10]}', str(days(bear_s, str(r[2])[:10])), f'{float(r[3])*100:.2f}%'])

# P10 2016-2017 中证红利全收益 vs 中证全指全收益（Wind，数据不用更新）
print('取 P10 Wind...')
p10 = [('中证红利全收益', E.fetch_kline('H00922.CSI','2016-01-04','2017-12-29')),
       ('中证全指全收益', E.fetch_kline('H00985.CSI','2016-01-04','2017-12-29'))]

print('数据读取 P2',[len(s) for _,s in p2],'P8行',len(p8_rows),'P10',[len(s) for _,s in p10],'P12',[len(s) for _,s in p12])

PEND = '⚠️ 此页为内部小程序截图/特殊数据，待提供后补入'

def build():
    prs = E.Presentation(E.TMPL_PATH)
    body = E.get_layout(prs, 1, '自定义版式'); cover = E.get_layout(prs, 4, '标题幻灯片')

    def placeholder(title, note):
        sl = prs.slides.add_slide(body)
        for ph in list(sl.placeholders): ph.element.getparent().remove(ph.element)
        E.add_tb(sl, title, E.CONTENT_L, E.CONTENT_TOP, E.CONTENT_W,
                 E.text_height_cm(title,E.TITLE_PT,E.CONTENT_W)+0.2, size=E.title_fit(title),
                 bold=True, color=E.C_BLACK, align=E.PP_ALIGN.CENTER)
        E.add_tb(sl, note, E.CONTENT_L+3, 8.5, E.CONTENT_W-6, 2, size=24, color=E.C_GRAY, align=E.PP_ALIGN.CENTER)

    def multi(title, pairs, left_table=False, points=None):
        if left_table: E.multi_chart_table_page(prs, body, title, msl(pairs), y_pct=True, points=points)
        else: E.multi_chart_page(prs, body, title, msl(pairs), points=points, y_pct=True)

    # P1
    E.cover_page(prs, cover, '这轮牛市跟历史比较，估值在什么位置？')
    # P2
    multi('近一年多涨幅：A股港股领涨全球', p2)
    # P3
    E.box_page(prs, body, '这一轮牛市上涨的原因',
               [('美元降息周期','美元利率、汇率贬值'),('股市基本面复苏','盈利同比恢复增长')],
               '这轮牛市能延续多久，也主要看上面两个因素。')
    # P4 双轴 = 两个原生图表叠放：底层面积(国债收益率,Y2右轴%) + 顶层透明折线(美元指数,Y1左轴)
    ym={x['date']:x['close'] for x in p4_yld}
    um={x['date']:x['close'] for x in p4_usd}
    cdates=[x['date'] for x in p4_usd if x['date'] in ym]
    idx_of={d:i for i,d in enumerate(cdates)}
    hbp={m['date']:m['bp'] for m in p4_marks if m['type']=='hike'}
    cbp={m['date']:m['bp'] for m in p4_marks if m['type']=='cut'}
    hike_vals=[um[d] if d in hbp else None for d in cdates]
    cut_vals =[um[d] if d in cbp else None for d in cdates]
    hlab={idx_of[d]:bp for d,bp in hbp.items() if d in idx_of}
    clab={idx_of[d]:bp for d,bp in cbp.items() if d in idx_of}
    sl=prs.slides.add_slide(body)
    for ph in list(sl.placeholders): ph.element.getparent().remove(ph.element)
    _t4='原因一：美元降息周期，美元利率汇率贬值'
    E.add_tb(sl,_t4,E.CONTENT_L,E.CONTENT_TOP,E.CONTENT_W,1.2,
             size=E.title_fit(_t4),bold=True,color=E.C_BLACK,align=E.PP_ALIGN.CENTER)
    ct=E.CONTENT_TOP+1.8; chh=E.CONTENT_BOTTOM-0.6-ct; cw=min(E.CONTENT_W,chh*2.5); cl=E.CONTENT_L+(E.CONTENT_W-cw)/2
    E.add_overlay_dual_axis(sl,cl,ct,cw,chh,cdates,
        '美元指数（衡量美元汇率变化）',[um[d] for d in cdates],
        '美股10年期国债收益率',[ym[d] for d in cdates],
        marker_series=[('加息',hike_vals,'FF2700',hlab),('降息',cut_vals,'2BA84A',clab)])
    # P5
    E.chart_below_table_page(prs, body, '原因二：股市基本面复苏，盈利同比恢复增长', msl(p5),
                             p5_table['headers'], p5_table['rows'], y_pct=True)
    # P6
    E.chart_page(prs, body, '这轮牛市，涨到什么水平了？', p6_line, show_hilo=False,
                 points=['中证全指点数，超过2021年最高点。','这轮牛市中证全指的上涨幅度，也超过上轮牛市涨幅。'],
                 bands=[{'start':'2018-10-19','end':'2021-02-28','label':'2021年牛市 +76.47%'},
                        {'start':'2024-09-18','end':'2026-06-06','label':'本轮牛市 +81.50%'}])
    # P7 星级 占位
    placeholder('这轮牛市，涨到什么水平了？',
                '（2）估值达到3.5星，跟2021年最高估值接近。\n' + PEND + '：中证全指星级走势(2012至今)')
    # P8
    E.summary_table_page.__wrapped__ if False else None
    _build_table(prs, body, '过去20年，A股经历的6轮牛熊市', p8_rows)
    # P9
    multi('2014-2015年牛市：中小盘成长风格领涨', p9)
    # P10
    multi('2016-2017年牛市：大盘价值风格领涨', p10)
    # P11 白酒 占位
    placeholder('2020-2021年牛市：大盘成长风格领涨', PEND + '：中证白酒全收益（Wind代码待确认）')
    # P12
    multi('2025-2026年牛市：小盘成长股领涨', p12)
    # P13 盈利相似（复用P5）
    E.chart_below_table_page(prs, body, '这轮牛市，和2013-2017年那一轮比较相似', msl(p5),
                             p5_table['headers'], p5_table['rows'], y_pct=True,
                             points=['（1）上市公司盈利情况相似：都是从低迷转向复苏。'])
    # P14 信号板 占位
    placeholder('这轮牛市，和2013-2017年那一轮比较相似',
                '（2）市场消息相似：都推出了一揽子的降息降准等利好政策，推动市场上涨。\n' + PEND + '：股市牛熊信号板截图')
    # P15
    multi('这轮牛市，和2013-2017年那一轮比较相似', p15,
          points=['（3）第一波上涨的品种类似：都是证券保险。'])
    # P16
    multi('这轮牛市，和2013-2017年那一轮比较相似', p16, left_table=True,
          points=['（4）不同品种的表现相似：中小盘、科技等成长风格表现突出，价值风格、消费等涨幅相对一般。'])
    # P17 止盈图 占位
    placeholder('3点几星，该如何投资？',
                '（1）跌出机会就买入，涨出机会就卖出，其他时间耐心等待。\n' + PEND + '：跌买涨卖示意图')
    # P18
    multi('3点几星，该如何投资？', p18,
          points=['（2）存量资金、止盈出来的资金，放在稳健策略中，例如365天、90天等组合。',
                  '最近指增组合止盈出来的资金，也放到365天组合里了。'])
    # P19
    E.box_page(prs, body, '3点几星，做定投需要注意什么？',
               [('①','不太适合一次性投入大笔资金到股票基金中'),('②','低估的品种，做少量资金定投是OK的'),
                ('③','做好长期定投计划，确保回到4-5星时有资金加仓')])
    # P20 估值表 占位
    placeholder('3点几星，做定投需要注意什么？', PEND + '：螺丝钉估值表截图（每晚发文/今天几星）')
    # P21
    E.arrows_page(prs, body, '总结', [
        '最近1年多，A股港股领涨全球。',
        '从市场整体来看，中证全指点数已超过2021年最高点，估值跟2021年最高估值接近。',
        '对比历史，这一轮牛市和2013-2017年那一轮比较类似。',
        '从具体品种看，这轮牛市小盘成长风格遥遥领先，红利等品种还有一部分在低估。',
        '当下3点几星，不宜大笔一次性投入；低估品种可少量定投；做好长期计划，回到4-5星时有资金加仓。'])

    for _ in range(8): E.remove_slide(prs, 0)
    out=os.path.expanduser('~/Downloads/449期直播PPT.pptx'); prs.save(out)
    print('\n✅ 完成！%d 张 → %s' % (len(prs.slides), out))

def _build_table(prs, body, title, rows):
    """直接用现成行数据建6轮牛熊表（不再Wind取数）。"""
    from pptx.util import Cm, Pt
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.dml.color import RGBColor
    sl = prs.slides.add_slide(body)
    for ph in list(sl.placeholders): ph.element.getparent().remove(ph.element)
    E.add_tb(sl, title, E.CONTENT_L, E.CONTENT_TOP, E.CONTENT_W,
             E.text_height_cm(title,E.TITLE_PT,E.CONTENT_W)+0.2, size=E.title_fit(title),
             bold=True, color=E.C_BLACK, align=PP_ALIGN.CENTER)
    headers=['轮次','牛/熊','起止时间','持续天数','最大涨跌幅']
    nr=len(rows)+1; top=E.CONTENT_TOP+1.4; tbl_h=E.CONTENT_BOTTOM-top-0.3
    g=sl.shapes.add_table(nr,5,Cm(E.CONTENT_L),Cm(top),Cm(E.CONTENT_W),Cm(tbl_h)).table
    for c,w in enumerate([3.6,2.6,13.5,5.0,7.3]): g.columns[c].width=Cm(w)
    rh=min(1.05,tbl_h/nr)
    for r in range(nr): g.rows[r].height=Cm(rh)
    def cell(r,c,t,bold=False,color=E.C_DARK,fill=None,align=PP_ALIGN.CENTER):
        cl=g.cell(r,c); cl.fill.solid(); cl.fill.fore_color.rgb=fill or RGBColor(0xFF,0xFF,0xFF)
        cl.margin_top=Cm(0.02); cl.margin_bottom=Cm(0.02); cl.vertical_anchor=MSO_ANCHOR.MIDDLE
        tf=cl.text_frame; tf.word_wrap=False; p=tf.paragraphs[0]; p.alignment=align
        run=p.add_run(); run.text=t; run.font.size=Pt(18); run.font.bold=bold; run.font.color.rgb=color; E.set_run_font(run)
    for c,h in enumerate(headers): cell(0,c,h,True,RGBColor(0xFF,0xFF,0xFF),E.C_MAIN)
    for r,row in enumerate(rows,1):
        bg=RGBColor(0xF5,0xF5,0xF5) if r%2==0 else RGBColor(0xFF,0xFF,0xFF)
        for c,v in enumerate(row):
            cell(r,c,v,bold=(c==0),color=(E.C_MAIN if c==4 else E.C_DARK),fill=bg,align=(PP_ALIGN.LEFT if c==2 else PP_ALIGN.CENTER))

build()
