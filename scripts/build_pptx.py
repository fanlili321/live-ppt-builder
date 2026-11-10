#!/usr/bin/env python3
# 465全本 → pptx(转key前)。页框用背景图,标题/正文/图表/表格为原生可编辑对象。
import json, os
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_TICK_MARK
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

W=os.path.dirname(os.path.abspath(__file__)); REUSE=os.path.join(W,'reuse'); FR=os.path.join(W,'frames')
E=lambda px:Emu(int(px*9525))
ORANGE=RGBColor(0xF1,0x76,0x40); GRAY=RGBColor(0x92,0x92,0x92); BLACK=RGBColor(0x0F,0x14,0x23)
SRC=RGBColor(0x4C,0x4C,0x4C); GRID=RGBColor(0xE0,0xE0,0xE0); WHITE=RGBColor(0xFF,0xFF,0xFF)
HDR=RGBColor(0xF1,0x76,0x40); GCOL=RGBColor(0xEF,0xEF,0xEF); MI=RGBColor(0xFD,0xF6,0xEE)  # 表头统一橙;首列/斑马灰;斑马米
FB="思源黑体 CN"

def load(n): return json.load(open(os.path.join(W,'data',n)))
def cum(rows,s=None,e=None):
    sel=[r for r in rows if (s is None or r['date']>=s) and (e is None or r['date']<=e)]
    b=sel[0]['close']; return [(r['date'],round((r['close']/b-1)*100,2)) for r in sel]
def samp(series,n):
    if len(series)<=n: return series
    st=(len(series)-1)/(n-1); idx=sorted(set([round(i*st) for i in range(n)]+[len(series)-1])); return [series[i] for i in idx]
def sparse(series,keep=7,md=False):
    n=len(series); step=max(1,(n-1)//(keep-1)); out=[]
    for i,(d,_) in enumerate(series): out.append((d[5:10] if md else d[:7]) if (i%step==0 or i==n-1) else "")
    return out
def yend(rows,y):
    ys=[r for r in rows if r['date'][:4]==str(y)]; return ys[-1]['close'] if ys else rows[-1]['close']

idx=load("p2p3_indices.json"); hl=idx['hongli_lowvol_TR']['rows']; kc=idx['kechuang50_TR']['rows']
hongli=load("p5p18_hongli_TR.json")['rows']; LATEST=hl[-1]['date']

prs=Presentation(); prs.slide_width=E(1920); prs.slide_height=E(1080); B=prs.slide_layouts[6]

def bg(sl,img): sl.shapes.add_picture(os.path.join(FR,img),0,0,E(1920),E(1080))
def full(sl,path): sl.shapes.add_picture(path,0,0,E(1920),E(1080))
def txt(sl,s,x,y,w,h,size,color,bold=False,align=PP_ALIGN.LEFT,font=FB):
    tb=sl.shapes.add_textbox(E(x),E(y),E(w),E(h)); tf=tb.text_frame; tf.word_wrap=True
    for i,line in enumerate(s.split("\n")):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        p.alignment=align; r=p.add_run(); r.text=line
        r.font.size=Pt(size); r.font.bold=bold; r.font.color.rgb=color; r.font.name=font
    return tb
def title(sl,s):
    n=len(s)  # 自适应:短标题90,长标题缩到一行不换行(模板规律)
    size=90 if n<=12 else (76 if n<=16 else (64 if n<=20 else 56))
    tb=txt(sl,s,90,120,1740,140,size,BLACK,bold=True,align=PP_ALIGN.CENTER)
    tb.text_frame.word_wrap=False
def source(sl,s): txt(sl,s,1180,1018,600,45,30,SRC,align=PP_ALIGN.RIGHT)  # 右端~1780,给页码让位
from pptx.enum.shapes import MSO_SHAPE
NUMG=RGBColor(0x99,0x99,0x99)
def pagenum(sl,n,cover_old=False):
    if cover_old:  # 复用截图:白盖右下角旧页码
        r=sl.shapes.add_shape(MSO_SHAPE.RECTANGLE,E(1832),E(1006),E(88),E(70))
        r.fill.solid(); r.fill.fore_color.rgb=WHITE; r.line.fill.background()
    txt(sl,str(n),1815,1016,90,55,38,NUMG,align=PP_ALIGN.RIGHT)

def style_chart(ch,colors):
    ch.has_legend=True; ch.legend.position=XL_LEGEND_POSITION.TOP; ch.legend.include_in_layout=False
    ch.legend.font.size=Pt(20); ch.legend.font.name=FB
    for i,se in enumerate(ch.series):
        se.format.line.color.rgb=colors[i]; se.format.line.width=Pt(2.5); se.smooth=False
    ca=ch.category_axis; va=ch.value_axis
    ca.has_major_gridlines=False; va.has_major_gridlines=True
    va.major_gridlines.format.line.color.rgb=GRID; va.major_gridlines.format.line.width=Pt(0.75)
    ca.format.line.color.rgb=GRID
    ca.tick_labels.font.size=Pt(16); ca.tick_labels.font.name=FB
    va.tick_labels.font.size=Pt(16); va.tick_labels.font.name=FB
    ca.major_tick_mark=XL_TICK_MARK.NONE; va.major_tick_mark=XL_TICK_MARK.NONE
def lchart(sl,x,y,w,h,cats,series,colors):
    cd=CategoryChartData(); cd.categories=cats
    for nm,vals in series: cd.add_series(nm,vals)
    gf=sl.shapes.add_chart(XL_CHART_TYPE.LINE,E(x),E(y),E(w),E(h),cd); style_chart(gf.chart,colors); return gf.chart

# ---------------- 逐页 ----------------
# P1 封面
s=prs.slides.add_slide(B); bg(s,'cover_bg.png')
txt(s,"红利指数涨起来了，\n还能投吗？",120,300,1000,300,72,ORANGE,bold=True)
# P2
s=prs.slides.add_slide(B); bg(s,'content_bg.png'); title(s,"近期，红利品种表现不错")
txt(s,"红利指数属于价值风格，遇成长风格回调时表现较好（如7月以来）；遇科创50等成长强势时（如4-6月），会跑输市场。",110,285,1700,80,32,BLACK)
L_h=samp(cum(hl,'2026-04-01','2026-06-30'),40); L_k=samp(cum(kc,'2026-04-01','2026-06-30'),40)
R_h=samp(cum(hl,'2026-07-01',LATEST),25); R_k=samp(cum(kc,'2026-07-01',LATEST),25)
lchart(s,80,412,860,528,sparse(L_h,7,True),[("红利低波全收益",[v for _,v in L_h]),("科创50全收益",[v for _,v in L_k])],[ORANGE,GRAY])
lchart(s,980,412,860,528,sparse(R_h,6,True),[("红利低波全收益",[v for _,v in R_h]),("科创50全收益",[v for _,v in R_k])],[ORANGE,GRAY])
source(s,"数据来源：Wind")
# P3
s=prs.slides.add_slide(B); bg(s,'content_bg.png'); title(s,"拉长时间来看，成长、价值风格的长期表现相差不大")
p3h=samp(cum(hl),50); p3k=samp(cum(kc),50)
lchart(s,110,300,1700,415,sparse(p3h,7),[("红利低波全收益",[v for _,v in p3h]),("科创50全收益",[v for _,v in p3k])],[ORANGE,GRAY])
h26=round((hl[-1]['close']/yend(hl,2025)-1)*100,2); k26=round((kc[-1]['close']/yend(kc,2025)-1)*100,2)
yrs=['指数名称','2020年','2021年','2022年','2023年','2024年','2025年','2026年至今']
row_h=['红利低波全收益','1.99%','16.34%','3.46%','12.23%','24.82%','5.57%',f'{h26}%']
row_k=['科创50全收益','40.02%','0.65%','-31.08%','-10.61%','16.89%','36.52%',f'{k26}%']
from pptx.oxml.ns import qn
gf3=s.shapes.add_table(3,8,E(110),E(752),E(1700),E(258)); tb=gf3.table
# 去边框(No Style No Grid)+关自带斑马,自己上色块
sid=tb._tbl.tblPr.find(qn('a:tableStyleId'))
if sid is not None: sid.text='{2D5ABB26-0587-4C30-8999-92F81FD0307C}'
tb.first_row=False; tb.horz_banding=False
tb.columns[0].width=E(360)
for j in range(1,8): tb.columns[j].width=E((1700-360)//7)
for rr in tb.rows: rr.height=E(86)
CREAM=RGBColor(0xFD,0xF6,0xEE)
def cellset(c,v,size,color,bold,fill):
    c.fill.solid(); c.fill.fore_color.rgb=fill; c.vertical_anchor=MSO_ANCHOR.MIDDLE
    c.margin_top=E(2); c.margin_bottom=E(2)
    p=c.text_frame.paragraphs[0]; p.alignment=PP_ALIGN.CENTER
    r=p.add_run(); r.text=v; r.font.size=Pt(size); r.font.bold=bold; r.font.color.rgb=color; r.font.name=FB
for j,v in enumerate(yrs): cellset(tb.cell(0,j),v,24,WHITE,True,HDR)  # 表头橙块白字
for ri,row in enumerate([row_h,row_k],start=1):
    rf=CREAM if ri==1 else WHITE  # 极浅斑马:红利米、科创白
    for j,v in enumerate(row):
        cellset(tb.cell(ri,j),v,26,BLACK,(j==0),rf)  # 首列白底加粗居中
source(s,"数据来源：Wind")
# P4
s=prs.slides.add_slide(B); full(s,os.path.join(REUSE,'p4_from_458p14.jpeg'))
# P5
s=prs.slides.add_slide(B); bg(s,'content_bg.png'); title(s,"红利品种的投资方法（2）持股收息")
p5=samp(cum(hongli),50)
lchart(s,110,300,1700,480,sparse(p5,8),[("某红利指数基金的表现（累计涨幅%）",[v for _,v in p5])],[ORANGE])
txt(s,"红利指数基金的基金分红，来自于股息分红。股息分红来自上市公司盈利的一部分，不受牛熊市涨跌影响。",470,830,1000,110,30,BLACK,align=PP_ALIGN.CENTER)
source(s,f"数据来源：Wind（截至{LATEST}）")
# P6-P12 复用图
for img in ['p6_from_463p7.jpeg','p7_from_439p6.jpeg','p8_股息率.png','p9_盈利收益率.png','0705_p23-23.jpg','p11_市净率.png','0705_p35-35.jpg']:
    s=prs.slides.add_slide(B); full(s,os.path.join(REUSE,img))
# P13
s=prs.slides.add_slide(B); bg(s,'content_bg.png'); title(s,"低估的红利品种，当前可以投资吗？")
txt(s,"• 红利星级整体在4点几星级，也在低估。",200,340,1520,70,34,BLACK)
txt(s,"• 一些低估的红利指数，可以考虑定投，逐渐积累股票资产；同时做好长期定投计划，确保回到5星级时仍有资金加仓。",200,450,1520,140,34,BLACK)
txt(s,"• 如果有一大笔资金，一次性大比例投入，更建议耐心等待5星级。",200,640,1520,70,34,BLACK)
# P14
s=prs.slides.add_slide(B); full(s,os.path.join(REUSE,'0705_p33-33.jpg'))
# P15 图+气泡
s=prs.slides.add_slide(B); full(s,os.path.join(REUSE,'0705_p34-34.jpg'))
txt(s,"参考螺丝钉红利星级及估值表，判断便宜还是贵。",560,905,820,70,26,ORANGE,bold=True,align=PP_ALIGN.CENTER)
# P16 图+改文字
s=prs.slides.add_slide(B); full(s,os.path.join(REUSE,'0705_p21-21.jpg'))
s.shapes.add_picture(os.path.join(REUSE,'white_cover.png'),E(60),E(298),E(1800),E(100))
txt(s,"所以历史上红利指数，高估比较少见，主要出现在2007年、2015年大牛市。其他时间里，多为低估和正常估值。",90,312,1740,90,26,BLACK)
# P17
s=prs.slides.add_slide(B); full(s,os.path.join(REUSE,'0705_p36-36.jpg'))
# P18
s=prs.slides.add_slide(B); bg(s,'content_bg.png'); title(s,"某红利指数基金的历次分红")
lchart(s,110,300,1700,480,sparse(p5,8),[("某红利指数基金的表现（累计涨幅%）",[v for _,v in p5])],[ORANGE])
txt(s,"红利指数基金的基金分红，来自于股息分红。股息分红来自上市公司盈利的一部分，不受牛熊市涨跌影响。",470,830,1000,110,30,BLACK,align=PP_ALIGN.CENTER)
source(s,f"数据来源：Wind（截至{LATEST}）")
# P19 总结
s=prs.slides.add_slide(B); bg(s,'content_bg.png'); title(s,"总结")
txt(s,"近期红利品种表现较好，主要与市场风格轮动有关。拉长时间看，成长与价值风格的长期表现相差不大。",160,320,1600,90,32,BLACK)
txt(s,"投资红利品种，适合低估时买入、持股收息。",160,440,1600,60,32,BLACK)
txt(s,"判断估值时，可以股息率为主，结合盈利收益率、市盈率和市净率综合判断，也可参考螺丝钉红利星级及估值表。",160,540,1600,90,32,BLACK)
txt(s,"红利指数基金既可以根据估值止盈，也可以长期持有，通过分红获得现金流。",160,660,1600,80,32,BLACK)

# ---- 统一重新生成页码(封面P1不编号;复用截图页先白盖旧页码) ----
reuse_idx={4,6,7,8,9,10,11,12,14,15,16,17}
for i,sl in enumerate(prs.slides, start=1):
    if i==1: continue
    pagenum(sl,i,cover_old=(i in reuse_idx))
prs.save(os.path.join(W,'465红利_claude输出.pptx')); print("full pptx已生成,页数=",len(prs.slides._sldIdLst))
