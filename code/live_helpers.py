# -*- coding: utf-8 -*-
"""直播PPT共享页型库(从450验证版抽取)。各期 build_xxx_mine.py 复用。
规则见记忆 feedback-ppt453-chart-rules(A–N) + 设计素材库。"""
import openpyxl, warnings
import ppt_engine as E
from pptx.util import Cm as _Cm, Pt as _Pt, Emu
def Cm(v): return _Cm(v*E.SCALE)   # 与引擎一致:输出层×SCALE放大到67.7
def Pt(v): return _Pt(v*E.SCALE)
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
warnings.filterwarnings('ignore')

# ── 素材库色板(设计素材库.md / reference-ppt-design-assets) ──
ORG       =RGBColor(0xF1,0x76,0x40)  # 主色橙
BLUE      =RGBColor(0x48,0x74,0xCB)  # 主蓝
YELLOW    =RGBColor(0xF2,0xBA,0x02)  # 主黄
GREEN     =RGBColor(0x75,0xBD,0x42)  # 主绿
RED       =RGBColor(0xE5,0x4C,0x5E)  # 主红
# 气泡/对比盒(取自模板pptx真实色)
PALE      =RGBColor(0xFF,0xF4,0xEA)  # 浅橙底(气泡/盒/节点)  ← 模板真实色
PALE_EDGE =RGBColor(0xF3,0xC7,0xA7)  # 浅橙描边(气泡/盒，柔和)  ← 模板真实色
PALE_BLUE =RGBColor(0xDC,0xE7,0xF8)  # 浅蓝底
BLUE_BUB  =RGBColor(0x7F,0xA8,0xDA)  # 蓝气泡(白字)
LIGHTGRAY =RGBColor(0xEE,0xEE,0xEE)  # 浅灰底  ← 模板真实色
FRAME_YEL =RGBColor(0xFF,0xC0,0x00)  # 走势图黄框  ← 模板真实色
ROW_CREAM =RGBColor(0xFD,0xF6,0xEE)  # 表格米色行
GOLD      =RGBColor(0xF2,0xAE,0x52)  # 树状根圆金橙
BGRAY     =RGBColor(0x6B,0x7A,0xA6); LBLUE=RGBColor(0x9D,0xC3,0xE6)
# 估值三色
VAL_LOW=RGBColor(0x5F,0xA8,0x3C); VAL_MID=RGBColor(0xF0,0x8A,0x3C); VAL_HIGH=RGBColor(0xE2,0x4B,0x3B); VAL_BOND=RGBColor(0xBB,0xD9,0xF5)
COMBO='/Users/fanlili/Downloads/投顾组合净值数据_20200101_20260615.xlsx'
_ws=openpyxl.load_workbook(COMBO,read_only=True,data_only=True).active

def combo(name,col,start=None):
    out=[]
    for r in _ws.iter_rows(min_row=2,values_only=True):
        if r[1]==name and r[col] not in (None,''):
            d=str(r[0])[:10]
            if start and d<start: continue
            try: out.append({'date':d,'close':float(r[col])})
            except: pass
    out.sort(key=lambda x:x['date']); return out

def kl(code,start,end,freq):
    y0=int(start[:4]); y1=int(end[:4]); out=[]
    for y in range(y0,y1+1):
        s=f'{y}-01-01' if y>y0 else start; e=f'{y}-12-31' if y<y1 else end
        for _ in range(8):
            try: out.extend(E.fetch_kline(code,s,e,freq)); break
            except Exception: continue
    out.sort(key=lambda x:x['date']); return out

def pct(name,data):
    b=data[0]['close']; return {'name':name,'dates':[x['date'] for x in data],'values':[(x['close']/b-1)*100 for x in data]}

def newslide(prs,body,title,align=PP_ALIGN.CENTER,tw=None):
    sl=prs.slides.add_slide(body)
    for ph in list(sl.placeholders): ph.element.getparent().remove(ph.element)
    w=tw or E.CONTENT_W
    E.add_tb(sl,title,E.CONTENT_L,E.CONTENT_TOP,w,E.text_height_cm(title,E.TITLE_PT,w)+0.2,
             size=E.title_fit(title),bold=True,color=E.C_BLACK,align=align)
    return sl

def btop(title): return E.CONTENT_TOP+E.text_height_cm(title,E.TITLE_PT,E.CONTENT_W)+0.2+0.9

def section(prs,body,title,number=''):
    """过渡页：橙块+白圆角卡+大号橙编号(01/02)+黑标题(照人工成品)。"""
    sl=prs.slides.add_slide(body)
    for ph in list(sl.placeholders): ph.element.getparent().remove(ph.element)
    band=sl.shapes.add_shape(MSO_SHAPE.RECTANGLE,Cm(0),Cm(0),Cm(33.87),Cm(10.2))
    band.fill.solid(); band.fill.fore_color.rgb=RGBColor(0xE8,0x6A,0x35); band.line.fill.background(); band.shadow.inherit=False
    card=sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Cm(3.0),Cm(5.4),Cm(27.8),Cm(7.6))
    card.fill.solid(); card.fill.fore_color.rgb=RGBColor(0xFF,0xFF,0xFF); card.line.fill.background(); card.shadow.inherit=False
    ty=6.0
    if number:
        E.add_tb(sl,str(number),4.5,5.8,6.0,2.2,size=48,bold=True,color=hl_main(),align=PP_ALIGN.LEFT); ty=8.2
    tb=E.add_tb(sl,title,4.5,ty,24.8,3.0,size=40,bold=True,color=E.C_BLACK,align=PP_ALIGN.LEFT); tb.text_frame.vertical_anchor=MSO_ANCHOR.TOP

def hl_main(): return E.C_MAIN

def catalog(prs,body,items):
    """目录页：白卡 + 01/02/03 列表(照人工成品P2)。"""
    sl=prs.slides.add_slide(body)
    for ph in list(sl.placeholders): ph.element.getparent().remove(ph.element)
    band=sl.shapes.add_shape(MSO_SHAPE.RECTANGLE,Cm(0),Cm(0),Cm(33.87),Cm(7.5))
    band.fill.solid(); band.fill.fore_color.rgb=RGBColor(0xE8,0x6A,0x35); band.line.fill.background(); band.shadow.inherit=False
    card=sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Cm(3.0),Cm(3.2),Cm(27.8),Cm(12.0))
    card.fill.solid(); card.fill.fore_color.rgb=RGBColor(0xFF,0xFF,0xFF); card.line.fill.background(); card.shadow.inherit=False
    n=len(items); y0=4.6; gap=(12.0-1.6)/max(n,1)
    for i,it in enumerate(items):
        E.add_tb(sl,f'0{i+1}  {it}',5.2,y0+i*gap,23.0,1.4,size=26,bold=True,color=E.C_BLACK,align=PP_ALIGN.LEFT)

def _stats(s):
    """从pct序列算 年化/最大回撤/区间。"""
    vals=s['values']; dates=s['dates']
    tot=(100+vals[-1])/(100+vals[0])
    from datetime import date
    d0=date.fromisoformat(dates[0]); d1=date.fromisoformat(dates[-1]); yrs=max((d1-d0).days/365.0,0.5)
    ann=(tot**(1/yrs)-1)*100
    peak=vals[0]; mdd=0; ps=pe=dates[0]; cp=dates[0]
    for dt,v in zip(dates,vals):
        if v>peak: peak=v; cp=dt
        dd=(100+v)/(100+peak)-1
        if dd<mdd: mdd=dd; ps=cp; pe=dt
    return f'{ann:.2f}%', f'{mdd*100:.2f}%', f'{ps.replace("-","")}-{pe.replace("-","")}'

def ph_note(prs,body,title,note):
    sl=newslide(prs,body,title)
    box=sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Cm(E.CONTENT_L+2),Cm(8.0),Cm(E.CONTENT_W-4),Cm(4.5))
    box.fill.solid(); box.fill.fore_color.rgb=PALE; box.line.color.rgb=ORG; box.line.width=Pt(1.5); box.shadow.inherit=False
    tb=E.add_tb(sl,'⚠️ '+note,E.CONTENT_L+3,8.0,E.CONTENT_W-6,4.5,size=20,color=E.C_DARK,align=PP_ALIGN.LEFT); tb.text_frame.vertical_anchor=MSO_ANCHOR.MIDDLE

def bubble(sl,text,left,top,w):
    lines=E.estimate_lines(text,18,w-1.2); h=lines*E.line_height_cm(18)+0.8
    box=sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Cm(left),Cm(top),Cm(w),Cm(h))
    box.fill.solid(); box.fill.fore_color.rgb=PALE; box.line.color.rgb=PALE_EDGE; box.line.width=Pt(1.0); box.shadow.inherit=False  # 素材库:浅橙底+浅橙描边
    tb=E.add_tb(sl,text,left+0.6,top,w-1.2,h,size=18,color=E.C_DARK,align=PP_ALIGN.LEFT); tb.text_frame.vertical_anchor=MSO_ANCHOR.MIDDLE
    return top+h

def chart_pg(prs,body,title,subtitle,series,bub=None,stats=False):
    """N2/N6: 标题+(加粗副标题)+两侧留白走势图+图下气泡。stats=True 时图下加 年化/最大回撤 数据表(照人工成品)。"""
    sl=newslide(prs,body,title); y=btop(title)
    if subtitle:
        E.add_tb(sl,subtitle,E.CONTENT_L,y,E.CONTENT_W,0.9,size=22,bold=True,color=E.C_MAIN,align=PP_ALIGN.LEFT); y+=1.1
    MG=3.6; cl=E.CONTENT_L+MG; cw=E.CONTENT_W-2*MG
    bub_h=(E.estimate_lines(bub,18,cw-1.2)*E.line_height_cm(18)+0.8+0.3) if bub else 0
    st_h=(0.7*(len(series)+1)+0.4) if stats else 0
    ct=y+0.2; chh=E.CONTENT_BOTTOM-0.6-ct-bub_h-st_h
    E.add_multi_line_chart(sl,cl,ct,cw,chh,series,y_pct=True)
    E.add_tb(sl,'数据来源：Wind / 螺丝钉投顾组合净值',cl+cw-13,ct+chh+0.02,13,0.4,size=11,color=E.C_GRAY,align=PP_ALIGN.RIGHT)
    yy=ct+chh+0.45
    if stats: yy=_stats_table(sl,series,cl,yy,cw)+0.3
    if bub: bubble(sl,bub,cl,yy,cw)
    return sl

def _stats_table(sl,series,left,top,w):
    rows=len(series)+1; tot=w; gt=sl.shapes.add_table(rows,4,Cm(left),Cm(top),Cm(w),Cm(0.7*rows)).table
    ws=[w*0.34,w*0.22,w*0.22,w*0.22]
    for c,wd in enumerate(ws): gt.columns[c].width=Cm(wd)
    def cell(r,c,t,bold,color,fill):
        cl=gt.cell(r,c); cl.fill.solid(); cl.fill.fore_color.rgb=fill; cl.vertical_anchor=MSO_ANCHOR.MIDDLE
        p=cl.text_frame.paragraphs[0]; p.alignment=PP_ALIGN.CENTER
        run=p.add_run(); run.text=t; run.font.size=Pt(12); run.font.bold=bold; run.font.color.rgb=color; E.set_run_font(run)
    for c,h in enumerate(['','年化收益率','最大回撤','最大回撤出现时间']): cell(0,c,h,True,RGBColor(0xFF,0xFF,0xFF),E.C_MAIN)
    for r,s in enumerate(series,1):
        ann,mdd,rng=_stats(s); bg=RGBColor(0xF5,0xF5,0xF5) if r%2==0 else RGBColor(0xFF,0xFF,0xFF)
        for c,v in enumerate([s['name'],ann,mdd,rng]): cell(r,c,v,(c==0),E.C_DARK,bg)
    return top+0.7*rows

def left_right(prs,body,title,left_points,series,src,right_top_labels=None,extra=None):
    """N7: 左文字 + 右走势图。"""
    sl=prs.slides.add_slide(body)
    for ph in list(sl.placeholders): ph.element.getparent().remove(ph.element)
    E.add_tb(sl,title,E.CONTENT_L,E.CONTENT_TOP,E.CONTENT_W,E.text_height_cm(title,E.title_fit(title),E.CONTENT_W)+0.2,size=E.title_fit(title),bold=True,color=E.C_BLACK,align=PP_ALIGN.CENTER)
    ytop=btop(title)
    E.add_arrow_points(sl,left_points,ytop,pt=20,left=E.CONTENT_L,width=9.5)
    cl=11.0; cw=E.CONTENT_L+E.CONTENT_W-cl-0.3; ct=ytop; chh=E.CONTENT_BOTTOM-0.7-ct
    E.add_multi_line_chart(sl,cl,ct,cw,chh,series,y_pct=True)
    E.add_tb(sl,src,cl+cw-13,ct+chh+0.02,13,0.4,size=11,color=E.C_GRAY,align=PP_ALIGN.RIGHT)
    if right_top_labels:
        lx=cl+1.0
        for txt,col in right_top_labels:
            E.add_tb(sl,txt,lx,ct-0.1,4.5,0.7,size=12,bold=True,color=col,align=PP_ALIGN.LEFT); lx+=5.0
    if extra: extra(sl,cl,ct,cw,chh)
    return sl

def smile(prs,body,title,subtitle=None):
    """微笑曲线：三色带+U形曲线+打点+右侧标注。"""
    sl=newslide(prs,body,title); y0=btop(title)+(1.0 if subtitle else 0)
    if subtitle: E.add_tb(sl,subtitle,E.CONTENT_L,btop(title),E.CONTENT_W,0.9,size=20,bold=True,color=E.C_MAIN,align=PP_ALIGN.LEFT)
    L=E.CONTENT_L+0.5; W=24.0; top=y0+0.2; H=E.CONTENT_BOTTOM-top-0.3
    bands=[(RGBColor(0xF6,0xC2,0x9B),'高估卖出'),(RGBColor(0xFC,0xE3,0xA8),'正常估值\n持有'),(RGBColor(0xBC,0xE4,0xC7),'低估买入')]
    bh=H/3
    for i,(c,lab) in enumerate(bands):
        by=top+i*bh
        b=sl.shapes.add_shape(MSO_SHAPE.RECTANGLE,Cm(L),Cm(by),Cm(W),Cm(bh))
        b.fill.solid(); b.fill.fore_color.rgb=c; b.line.fill.background(); b.shadow.inherit=False
        cap=sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Cm(L+W+0.4),Cm(by+bh/2-0.8),Cm(4.0),Cm(1.6))
        cap.fill.solid(); cap.fill.fore_color.rgb=PALE; cap.line.color.rgb=PALE_EDGE; cap.line.width=Pt(1.0); cap.shadow.inherit=False
        tb=E.add_tb(sl,lab,L+W+0.4,by+bh/2-0.8,4.0,1.6,size=16,bold=True,color=E.C_DARK,align=PP_ALIGN.CENTER); tb.text_frame.vertical_anchor=MSO_ANCHOR.MIDDLE
    ytop=top+0.6; ybot=top+H-0.8; pts=[]
    for k in range(41):
        t=k/40; x=L+1.0+(W-2.0)*t; yy=ybot-(ybot-ytop)*((2*t-1)**2); pts.append((x,yy))
    curve(sl,pts,ORG,3.0)
    for k in (8,13,18,20,22,27,32):
        x,yy=pts[k]
        dot=sl.shapes.add_shape(MSO_SHAPE.OVAL,Cm(x-0.18),Cm(yy-0.18),Cm(0.36),Cm(0.36))
        dot.fill.solid(); dot.fill.fore_color.rgb=RGBColor(0xFF,0xFF,0xFF); dot.line.color.rgb=ORG; dot.line.width=Pt(2); dot.shadow.inherit=False

def curve(sl,pts,color,wpt):
    U=360000
    fb=sl.shapes.build_freeform(Emu(int(pts[0][0]*U)),Emu(int(pts[0][1]*U)),scale=1)
    fb.add_line_segments([(Emu(int(x*U)),Emu(int(y*U))) for x,y in pts[1:]],close=False)
    shp=fb.convert_to_shape(); shp.fill.background(); shp.line.color.rgb=color; shp.line.width=Pt(wpt); shp.shadow.inherit=False
    return shp

def table_page(prs,body,title,headers,rows,widths,src='',fills=None):
    sl=newslide(prs,body,title)
    nr=len(rows)+1; t_h=E.text_height_cm('x',E.TITLE_PT,E.CONTENT_W)+0.2
    top=E.CONTENT_TOP+t_h+1.0; tbl_h=E.CONTENT_BOTTOM-top-0.9
    gt=sl.shapes.add_table(nr,len(headers),Cm(E.CONTENT_L),Cm(top),Cm(E.CONTENT_W),Cm(tbl_h)).table
    tot=sum(widths)
    for c,w in enumerate(widths): gt.columns[c].width=Cm(w/tot*E.CONTENT_W)
    for r in range(nr): gt.rows[r].height=Cm(tbl_h/nr)
    def cell(r,c,t,bold,color,fill,al):
        cl=gt.cell(r,c); cl.fill.solid(); cl.fill.fore_color.rgb=fill; cl.vertical_anchor=MSO_ANCHOR.MIDDLE
        p=cl.text_frame.paragraphs[0]; p.alignment=al; cl.text_frame.word_wrap=True
        run=p.add_run(); run.text=str(t); run.font.size=Pt(16); run.font.bold=bold; run.font.color.rgb=color; E.set_run_font(run)
    for c,h in enumerate(headers): cell(0,c,h,True,RGBColor(0xFF,0xFF,0xFF),E.C_MAIN,PP_ALIGN.CENTER)
    for r,row in enumerate(rows,1):
        bg=RGBColor(0xF5,0xF5,0xF5) if r%2==0 else RGBColor(0xFF,0xFF,0xFF)
        for c,v in enumerate(row): cell(r,c,v,(c==0),E.C_DARK,bg,(PP_ALIGN.LEFT if c==1 else PP_ALIGN.CENTER))
    if src: E.add_tb(sl,src,E.CONTENT_L+E.CONTENT_W-12,top+tbl_h+0.15,12,0.5,size=12,color=E.C_GRAY,align=PP_ALIGN.RIGHT)
    return sl

def pros_cons(prs,body,title,intro,pros,cons,bub=None):
    """优缺点双色块(照黄金455 P5/P6)：左橙'优点'卡+浅橙框，右蓝'缺点'卡+浅蓝框，可选底部气泡。"""
    sl=newslide(prs,body,title); y=btop(title)-0.3
    if intro:
        E.add_tb(sl,intro,E.CONTENT_L,y,E.CONTENT_W,0.9,size=20,color=E.C_DARK,align=PP_ALIGN.LEFT); y+=1.2
    bw=14.6; gap=2.0; bx=[E.CONTENT_L+0.6,E.CONTENT_L+0.6+bw+gap]
    fh=5.2; top=y+0.9
    for i,(lab,txt,lc,fc,ec) in enumerate([('优点',pros,ORG,PALE,PALE_EDGE),('缺点',cons,RGBColor(0x5B,0x6B,0x9E),PALE_BLUE,BLUE_BUB)]):
        x=bx[i]
        frame=sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Cm(x),Cm(top),Cm(bw),Cm(fh))
        frame.fill.solid(); frame.fill.fore_color.rgb=fc; frame.line.color.rgb=ec; frame.line.width=Pt(1.3); frame.shadow.inherit=False
        tag=sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Cm(x+0.6),Cm(top-0.8),Cm(3.4),Cm(1.5))
        tag.fill.solid(); tag.fill.fore_color.rgb=lc; tag.line.fill.background(); tag.shadow.inherit=False
        tt=E.add_tb(sl,lab,x+0.6,top-0.8,3.4,1.5,size=18,bold=True,color=RGBColor(0xFF,0xFF,0xFF),align=PP_ALIGN.CENTER); tt.text_frame.vertical_anchor=MSO_ANCHOR.MIDDLE
        bt=E.add_tb(sl,txt,x+0.8,top+0.9,bw-1.6,fh-1.2,size=18,color=E.C_DARK,align=PP_ALIGN.LEFT); bt.text_frame.vertical_anchor=MSO_ANCHOR.TOP
    if bub: bubble(sl,bub,E.CONTENT_L+2,top+fh+0.6,E.CONTENT_W-4)
    return sl

def cover(prs,cover_layout,title): E.cover_page(prs,cover_layout,title)
def summary(prs,body,title,points): E.arrows_page(prs,body,title,points,pt=24,spread=True)
def img(prs,body,title,path): E.img_page(prs,body,title,path)

def img_reuse(prs,body,img_path,title=None):
    """复用历史成品页：整页贴高清jpeg(同模板，视觉100%还原)，盖掉自带旧页码。
    页码由 paginate() 统一重编。仅用于'完全不改'或'仅改标题(传title)'的复用页；
    要改正文/删元素的页不能用(会残留旧内容)。"""
    sl=prs.slides.add_slide(body)
    for ph in list(sl.placeholders): ph.element.getparent().remove(ph.element)
    use_path=img_path
    # 仅改标题：用PIL把原图标题行涂白(shape在LibreOffice下不可靠)，再原生写新标题
    if title:
        import os
        from PIL import Image as _I, ImageDraw as _D
        im=_I.open(img_path).convert('RGB'); iw,ih=im.size
        _D.Draw(im).rectangle([0,int(ih*0.09),iw,int(ih*0.245)],fill=(255,255,255))  # 只盖主标题行,保副标题
        use_path='/tmp/'+os.path.basename(os.path.splitext(img_path)[0])+'_rt.jpg'; im.save(use_path,quality=92)
    sl.shapes.add_picture(use_path,0,0,prs.slide_width,prs.slide_height)
    W=prs.slide_width; H=prs.slide_height
    if title:
        E.add_tb(sl,title,E.CONTENT_L,2.2,E.CONTENT_W,2.0,size=E.title_fit(title),bold=True,color=E.C_BLACK,align=PP_ALIGN.CENTER)
    # 盖掉右下角自带旧页码(白块)
    cov=sl.shapes.add_shape(MSO_SHAPE.RECTANGLE,W-Cm(2.2),H-Cm(1.15),Cm(2.0),Cm(1.0))
    cov.fill.solid(); cov.fill.fore_color.rgb=RGBColor(0xFF,0xFF,0xFF); cov.line.fill.background(); cov.shadow.inherit=False
    return sl

def paginate(prs,skip_first=True):
    """全本统一重编页码：右下角灰字序号，封面(第1页)不显示。在 remove_slide 之后、save 之前调用。"""
    W=prs.slide_width; H=prs.slide_height
    for i,sl in enumerate(prs.slides):
        if skip_first and i==0: continue
        tb=sl.shapes.add_textbox(W-Cm(2.0),H-Cm(1.1),Cm(1.6),Cm(0.7)); tf=tb.text_frame
        tf.word_wrap=False; p=tf.paragraphs[0]; p.alignment=PP_ALIGN.RIGHT
        r=p.add_run(); r.text=str(i+1); r.font.size=Pt(12); r.font.color.rgb=E.C_GRAY; E.set_run_font(r)
