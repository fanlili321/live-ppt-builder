# -*- coding: utf-8 -*-
"""446期：财报更新，上市公司的盈利增长情况如何？
走势图(指数+全收益)原生取数；归母净利润柱状&同比表 需用户导出成分股数据→占位。复用417版式。"""
import sys,json; sys.path.insert(0,'/Users/fanlili/Downloads/live-ppt-tool')
import ppt_engine as E; import live_helpers as hl
OUT='/Users/fanlili/Desktop/范丽丽./直播PPT成品/446期直播PPT.pptx'
C=json.load(open('/Users/fanlili/Downloads/live-ppt-tool/cache_446.json'))

# 页标题(来自大纲，措辞待数据确认) 与 全收益名
TITLES={
 '中证全指':'盈利比较稳定','沪深300':'盈利稳定增长','恒生指数':'盈利比较稳定','H股指数':'盈利比较稳定',
 '中证500':'盈利波动较大','中证1000':'盈利波动较大','创业板指':'盈利波动较大','科创50':'盈利波动较大',
 '中证红利':'盈利稳定增长','中证红利低波':'盈利稳定增长','恒生红利低波动':'盈利波动较大','中证A500':'盈利比较稳定',
 '消费50':'盈利稳定增长','恒生消费':'盈利稳定增长','医药100':'盈利波动较大','恒生医疗保健':'盈利增长较快','港股科技':'盈利增长较快',
}
GROUPS=[('宽基指数',['中证全指','沪深300','恒生指数','H股指数','中证500','中证1000','创业板指','科创50']),
        ('策略指数',['中证红利','中证红利低波','恒生红利低波动','中证A500']),
        ('行业 / 主题指数',['消费50','恒生消费','医药100','恒生医疗保健','港股科技'])]
PH_BUB='⚠️ 归母净利润（柱状）及历年同比表（删2025Q1-Q3、新增2025全年+2026Q1），待提供该指数成分股归母净利润数据后原生补充。'

def idx_page(prs,body,name):
    rec=C.get(name,{}); title=f'{name}：{TITLES.get(name,"")}（待数据确认）'
    series=[]
    if rec.get('price'):
        series.append(hl.pct(name,rec['price']['data']))
        if rec.get('tr'): series.append(hl.pct(name+'全收益',rec['tr']['data']))
        hl.chart_pg(prs,body,title,'盈利长期增长，是指数上涨的主要动力',series,bub=PH_BUB)
    else:
        hl.ph_note(prs,body,title,f'走势图待生成：{name} 指数代码未取到数据，需确认代码后补。另需成分股归母净利润数据做柱状+同比表。')

def build():
    prs=E.Presentation(E.TMPL_PATH)
    body=E.get_layout(prs,1,'自定义版式'); cov=E.get_layout(prs,4,'标题幻灯片')
    hl.cover(prs,cov,'财报更新，\n上市公司的盈利增长情况如何？')
    hl.ph_note(prs,body,'上市公司的定期报告及披露时间','复用417期P2设计页，待提供参考图或原生重做（年报/半年报/季报披露时间）。')
    hl.ph_note(prs,body,'如何查询上市公司的定期报告？','复用417期P3设计页，待提供参考图或原生重做。')
    hl.ph_note(prs,body,'盈利长期增长，是市场上涨的主要动力','复用417期P4：走势图+归母净利润，需提供中证全指成分股归母净利润数据后原生生成。')
    hl.ph_note(prs,body,'不同类型的代表指数','417期P5设计页：宽基/策略/行业三类指数列表（创业板后补科创50、中证红利后补中证红利低波、删中证消费）。待原生重做。')
    for gname,members in GROUPS:
        hl.section(prs,body,gname)
        for nm in members: idx_page(prs,body,nm)
    hl.summary(prs,body,'总结（待数据更新）',[
        '2025年4季度，A股整体上市公司盈利同比增速，相比前几季度有所变化（待数据回填）。',
        '第一梯队：盈利增长最强、估值最高，代表为科技方向。',
        '第二梯队：盈利增长相对稳定，如红利、自由现金流等。',
        '第三梯队：地产和消费产业链，盈利同比增长下降。'])
    for _ in range(8): E.remove_slide(prs,0)
    prs.save(OUT); print(f'✅ 446完成！{len(prs.slides)} 张 → {OUT}')

build()
