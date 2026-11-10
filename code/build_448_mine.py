# -*- coding: utf-8 -*-
"""448期：存款利率下行，哪些策略可以作为「存款+」？（V2大纲）
原生：CPPI案例图(红利低波全收益+存款本息线)/90天组合/365天组合/三策略分隔页；
复用4.26活动/227/408/吴晓波/442 设计页+估值表+源文件 占位。"""
import sys,json; sys.path.insert(0,'/Users/fanlili/Downloads/live-ppt-tool')
import ppt_engine as E; import live_helpers as hl
OUT='/Users/fanlili/Desktop/范丽丽./直播PPT成品/448期直播PPT.pptx'
C=json.load(open('/Users/fanlili/Downloads/live-ppt-tool/cache_448.json'))

def dep_line(dates):
    """合成存款本息线：2021初100万，每年+1万到2026初105万。返回pct序列(相对100万)。"""
    vals=[]
    for dt in dates:
        y=int(dt[:4]); m=int(dt[5:7]); frac=(y-2021)+(m-1)/12.0
        vals.append({'date':dt,'close':100.0+frac})
    return vals

def build():
    prs=E.Presentation(E.TMPL_PATH)
    body=E.get_layout(prs,1,'自定义版式'); cov=E.get_layout(prs,4,'标题幻灯片')
    hl.cover(prs,cov,'存款利率下行，\n哪些策略可以作为「存款+」？')
    hl.ph_note(prs,body,'存款的收益低，并且在不断下降','复用4.26出版社线下活动P3，新增气泡"截至2026年3月底，居民存款余额超170万亿"。待提供参考图。')
    hl.ph_note(prs,body,'存款会搬家进股市吗？','复用227期P2，待提供参考图或原生重做。')
    hl.ph_note(prs,body,'有多少资金流入股市？','源文件 cunkan_plus_light.pptx（配色微调），待提供该文件。')
    hl.ph_note(prs,body,'当前「存款+」的趋势还在继续','源文件同上，待提供。')
    hl.ph_note(prs,body,'「存款+」策略有哪些？','三框导航：①短债策略 ②CPPI策略 ③「固收+」策略。待补三框组件后原生。')
    # 短债策略
    hl.section(prs,body,'短债策略')
    hl.ph_note(prs,body,'纯债基金的分类','复用408期P5，待提供参考图或原生重做。')
    hl.ph_note(prs,body,'短债基金和长债基金','复用408期P6，待提供参考图或原生重做。')
    hl.ph_note(prs,body,'影响纯债基金涨跌的主要因素——利率','复用408期P9，估值表图待提供。')
    hl.ph_note(prs,body,'螺丝钉人民币债券星级&估值表','复用408期P10，债券估值表图待提供。')
    # 90天组合(原生)
    d90=hl.pct('90天组合',C['d90']['data'])
    hl.chart_pg(prs,body,'90天组合：以短债基金为主','成立以来净值走势',[d90],
        bub='右下角"识别二维码查看、购买"气泡及二维码待提供。')
    hl.ph_note(prs,body,'90天组合收益来源','气泡：90天组合平均2天收到一个蛋；存款、货币基金平均3天收到一个蛋。（"收蛋"=0.01%收益/个）待原生配图。')
    # CPPI策略
    hl.section(prs,body,'CPPI策略')
    hl.ph_note(prs,body,'CPPI策略：原保本基金的常用策略','待提供参考图。')
    hl.ph_note(prs,body,'CPPI策略原理','源文件 资产配置.key 第22页，待提供该文件。')
    hl.ph_note(prs,body,'CPPI策略的案例','示意图，待提供参考图。')
    # CPPI案例图(原生)：红利低波全收益 + 存款本息线
    hlv=C['hongli_lowvol']['data']
    s1=hl.pct('中证红利低波全收益',hlv); s2=hl.pct('银行存款本息和',dep_line([x['date'] for x in hlv]))
    hl.chart_pg(prs,body,'CPPI策略的案例','5年后：100万存款本金 + 5万红利指数基金本金 + 红利指数投资收益',
        [s1,s2],bub='2021年至今：橙线=中证红利低波全收益，灰线=银行存款本息和(100万逐年+1万到105万)。利息投入低估红利指数，本金由存款保障。')
    hl.ph_note(prs,body,'CPPI策略的注意事项','两点框：①投资周期通常跟利息发放周期匹配，不适合短期；②利息投资方向适合低估的股票指数基金(如红利指数)。待补框组件。')
    # 固收+策略
    hl.section(prs,body,'「固收+」策略')
    hl.ph_note(prs,body,'什么是「固收+」','复用吴晓波频道20251222直播P21，待提供参考图。')
    hl.ph_note(prs,body,'「固收+」品种有哪些？','复用吴晓波P22，待提供参考图。')
    hl.ph_note(prs,body,'「固收+」品种的三个特点','复用吴晓波P23，待提供参考图。')
    hl.ph_note(prs,body,'特点一：股债负相关性，减少了「固收+」的波动风险','吴晓波P24，需股债指数数据后原生生成走势对比图。')
    hl.ph_note(prs,body,'特点二：收益和波动风险受股票比例影响','复用吴晓波P25，待提供参考图。')
    hl.ph_note(prs,body,'特点三：受益于存款利率下降','复用442期P2，数据待更新后原生生成。')
    hl.ph_note(prs,body,'月薪宝、365天投顾组合','复用442期P6，待提供参考图。')
    # 365天组合(原生)
    d365=hl.pct('365天组合',C['d365']['data'])
    hl.chart_pg(prs,body,'月薪宝、365天投顾组合','365天组合成立以来净值走势',[d365],
        bub='右下角"识别二维码查看、购买"气泡及二维码待提供。')
    hl.summary(prs,body,'总结',[
        '存款利率持续下行，在稳健基础上获得稍高收益的「存款+」策略受到欢迎。',
        '短债策略：波动风险较小，当前利息收益率在1.5%-1.7%。',
        'CPPI策略：俗称保本策略，大多数资金配存款/固收，少部分配股票，利息覆盖股票下跌波动。',
        '「固收+」策略：底层配置一定股票资产，收益和风险主要取决于股票比例。'])
    for _ in range(8): E.remove_slide(prs,0)
    prs.save(OUT); print(f'✅ 448完成！{len(prs.slides)} 张 → {OUT}')

build()
