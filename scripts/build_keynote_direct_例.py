#!/usr/bin/env python3
# 465期红利 直播Keynote 构建器（安全版）
# 安全规矩：只操作"文件路径以 465红利_claude输出.key 结尾"的那个文档；找不到即 error 中止；
#          不用 front document、不用 activate、不盲目重试变更、删除前核验页数。
import json, subprocess, os, sys, time

W = os.path.dirname(os.path.abspath(__file__))
REUSE = os.path.join(W, 'reuse')
MYNAME = "465红利_claude输出.key"   # 唯一标识

def esc(s):
    return s.replace('\\', '\\\\').replace('"', '\\"')

GUARD_HEAD = f'''tell application "Keynote"
with timeout of 600 seconds
set myDoc to missing value
repeat with dd in documents
try
if (file of dd as text) ends with "{MYNAME}" then set myDoc to dd
end try
end repeat
if myDoc is missing value then error "ABORT_NOT_OPEN: {MYNAME} 未打开，已中止（绝不误碰其它文档）"
tell myDoc
'''
GUARD_TAIL = '''end tell
end timeout
end tell'''

def osa(body, timeout=600):
    script = GUARD_HEAD + body + "\n" + GUARD_TAIL
    r = subprocess.run(['osascript', '-e', script], capture_output=True, text=True, timeout=timeout+30)
    if r.returncode != 0:
        print("OSA ERR（已中止，未继续）:", r.stderr.strip()[:400]); sys.exit(1)
    return r.stdout.strip()

def osa_app(body, timeout=120):
    # 需要在 application 作用域（非 tell myDoc）里做的只读/导出操作，仍先按路径锁定 myDoc
    script = f'''tell application "Keynote"
with timeout of {timeout} seconds
set myDoc to missing value
repeat with dd in documents
try
if (file of dd as text) ends with "{MYNAME}" then set myDoc to dd
end try
end repeat
if myDoc is missing value then error "ABORT_NOT_OPEN"
{body}
end timeout
end tell'''
    r = subprocess.run(['osascript', '-e', script], capture_output=True, text=True, timeout=timeout+30)
    if r.returncode != 0:
        print("OSA ERR:", r.stderr.strip()[:400]); sys.exit(1)
    return r.stdout.strip()

# ---------- 数据 ----------
def load(name): return json.load(open(os.path.join(W, 'data', name)))
def cum_series(rows, start=None, end=None):
    sel = [r for r in rows if (start is None or r['date'] >= start) and (end is None or r['date'] <= end)]
    base = sel[0]['close']
    return [(r['date'], round((r['close']/base-1)*100, 2)) for r in sel]
def sample(series, n):
    if len(series) <= n: return series
    step = (len(series)-1)/(n-1)
    idx = sorted(set([round(i*step) for i in range(n)] + [len(series)-1]))
    return [series[i] for i in idx]

idx = load('p2p3_indices.json')
hl = idx['hongli_lowvol_TR']['rows']; kc = idx['kechuang50_TR']['rows']
hongli = load('p5p18_hongli_TR.json')['rows']
LATEST = hl[-1]['date']
p2L_h = sample(cum_series(hl,'2026-04-01','2026-06-30'),30); p2L_k = sample(cum_series(kc,'2026-04-01','2026-06-30'),30)
p2R_h = sample(cum_series(hl,'2026-07-01',LATEST),20);        p2R_k = sample(cum_series(kc,'2026-07-01',LATEST),20)
p3_h = sample(cum_series(hl),40); p3_k = sample(cum_series(kc),40)
p5 = sample(cum_series(hongli),40)
def vals(s): return '{'+', '.join(str(v) for _,v in s)+'}'
def cols_sparse(s, keep=7, md=False):
    n=len(s); step=max(1,(n-1)//(keep-1))
    def lab(i,d): return (d[5:10] if md else d[:7]) if (i%step==0 or i==n-1) else ""
    return '{'+', '.join('"%s"'%lab(i,d) for i,(d,_) in enumerate(s))+'}'
def year_end_close(rows, year):
    ys=[r for r in rows if r['date'][:4]==str(year)]
    return ys[-1]['close'] if ys else rows[-1]['close']

# ---------- 原语 ----------
def add_image_slide(img):
    p = esc(os.path.join(REUSE, img))
    return f'set s to make new slide\ntell s\nmake new image with properties {{file:POSIX file "{p}", position:{{0, 0}}, width:1920, height:1080}}\nend tell'
def T(t,x,y,w,h,size=38,rgb=(0,0,0),font="SourceHanSansCN-Regular"):
    r,g,b=rgb
    return (f'set ti to make new text item with properties {{object text:"{esc(t)}", position:{{{x}, {y}}}, width:{w}, height:{h}}}\n'
            f'set size of object text of ti to {size}\n'
            f'set color of object text of ti to {{{r*257}, {g*257}, {b*257}}}\n'
            f'try\nset font of object text of ti to "{font}"\nend try')
# 页标题:思源黑体Bold 72 黑; 来源:思源黑体Regular 30 灰#4C4C4C
def TITLE(t): return T(t,110,120,1700,150,size=72,rgb=(0,0,0),font="SourceHanSansCN-Bold")
def SRC(t,x=1470,y=1035,w=430): return T(t,x,y,w,40,size=30,rgb=(76,76,76),font="SourceHanSansCN-Regular")
def add_chart(rownames, series, colseries, x, y, w, h, md=False):
    rn = '{'+', '.join('"%s"'%r for r in rownames)+'}'
    data = '{'+', '.join(vals(s) for s in series)+'}'
    return (f'tell s\nadd chart row names {rn} column names {cols_sparse(colseries, 7, md)} data {data} type line_2d group by chart row\nend tell\n'
            f'set c to last chart of s\nset position of c to {{{x}, {y}}}\nset width of c to {w}\nset height of c to {h}')

BRAND_TPL = "/Users/fanlili/Desktop/范丽丽./图/直播ppt/2025.8【指数系列ppt模版】.key"
def build_cover():
    # 跨文档：打开品牌模板，复制其封面(slide1)到myDoc末尾，关模板
    script = f'''tell application "Keynote"
with timeout of 600 seconds
set tdoc to open POSIX file "{BRAND_TPL}"
delay 3
set myDoc to missing value
repeat with dd in documents
try
if (file of dd as text) ends with "{MYNAME}" then set myDoc to dd
end try
end repeat
if myDoc is missing value then error "ABORT_COVER: myDoc未找到"
duplicate slide 1 of tdoc to after last slide of myDoc
close tdoc saving no
end timeout
end tell'''
    r = subprocess.run(['osascript','-e',script], capture_output=True, text=True, timeout=650)
    if r.returncode != 0:
        print("封面复制失败:", r.stderr.strip()[:300]); sys.exit(1)
    # 移到最前 + 改标题(把"北证50…"两处大标题换成465标题，删副标题)
    osa('''move slide (count of slides) to before slide 1
tell slide 1
repeat with ti in text items
set tx to object text of ti
if tx contains "北证50" then set object text of ti to "红利指数涨起来了，
还能投吗？"
if tx contains "对我们投资" then set object text of ti to ""
end repeat
end tell''')
    print("封面 ok")

def build_all():
    print("目标文件:", MYNAME)
    n0 = int(osa_app('return (count of slides of myDoc) as string'))
    print("品牌模板初始页数:", n0)
    if n0 < 2:
        print(f"异常：模板页数过少({n0})，中止。"); sys.exit(1)
    WHITE = esc(os.path.join(REUSE,'white_cover.png'))
    # P1 封面：品牌模板自带 slide1 是封面，只改标题
    osa('''tell slide 1
repeat with ti in text items
set tx to object text of ti
if tx contains "北证50" then set object text of ti to "红利指数涨起来了，
还能投吗？"
if tx contains "对我们投资" then set object text of ti to ""
end repeat
end tell'''); print("封面标题 ok")
    # P2
    osa(f'''set s to make new slide
tell s
{TITLE("近期，红利品种表现不错")}
{T("红利指数属于价值风格，遇成长风格回调时表现较好（如7月以来）；遇科创50等成长强势时（如4-6月），会跑输市场。",110,290,1700,80,size=32)}
{SRC("数据来源：Wind",1560,1035,340)}
end tell
{add_chart(["红利低波全收益","科创50全收益"],[p2L_h,p2L_k],p2L_h,90,380,860,560,md=True)}
{add_chart(["红利低波全收益","科创50全收益"],[p2R_h,p2R_k],p2R_h,980,380,860,560,md=True)}'''); print("P2 ok")
    # P3
    osa(f'''set s to make new slide
tell s
{TITLE("拉长时间来看，成长、价值风格的长期表现相差不大")}
{SRC("数据来源：Wind",1560,1035,340)}
end tell
{add_chart(["红利低波全收益","科创50全收益"],[p3_h,p3_k],p3_h,110,300,1700,440)}''')
    h2026=round((hl[-1]['close']/year_end_close(hl,2025)-1)*100,2)
    k2026=round((kc[-1]['close']/year_end_close(kc,2025)-1)*100,2)
    print(f"  P3 2026至今: 红利={h2026}% 科创={k2026}%")
    annual_h=[1.99,16.34,3.46,12.23,24.82,5.57,h2026]
    annual_k=[40.02,0.65,-31.08,-10.61,16.89,36.52,k2026]
    yrs=['2020年','2021年','2022年','2023年','2024年','2025年','2026年至今']
    tbl='set t to make new table with properties {row count:3, column count:8, position:{110, 815}, width:1700, height:205}\ntell t\ntry\nset width of column 1 to 340\nend try\nset value of cell 1 of row 1 to "指数名称"\n'
    tbl+='\n'.join(f'set value of cell {i+2} of row 1 to "{y}"' for i,y in enumerate(yrs))+'\n'
    tbl+='set value of cell 1 of row 2 to "红利低波全收益"\n'+'\n'.join(f'set value of cell {i+2} of row 2 to "{v}%"' for i,v in enumerate(annual_h))+'\n'
    tbl+='set value of cell 1 of row 3 to "科创50全收益"\n'+'\n'.join(f'set value of cell {i+2} of row 3 to "{v}%"' for i,v in enumerate(annual_k))+'\nend tell'
    osa(f'tell slide -1\n{tbl}\nend tell'); print("P3 ok")
    # P4
    osa(add_image_slide('p4_from_458p14.jpeg')); print("P4 ok")
    # P5
    osa(f'''set s to make new slide
tell s
{TITLE("红利品种的投资方法（2）持股收息")}
{T("红利指数基金的基金分红，来自于股息分红。股息分红来自上市公司盈利的一部分，不受牛熊市涨跌影响。",470,890,1000,110,size=30)}
{SRC("数据来源：Wind（截至"+LATEST+"）",1430,1030,480)}
end tell
{add_chart(["某红利指数基金的表现（累计涨幅%）"],[p5],p5,110,300,1700,500)}'''); print("P5 ok")
    # P6-P12 图
    for i,img in enumerate(['p6_from_463p7.jpeg','p7_from_439p6.jpeg','p8_股息率.png','p9_盈利收益率.png','0705_p23-23.jpg','p11_市净率.png','0705_p35-35.jpg']):
        osa(add_image_slide(img)); print(f"P{6+i} ok")
    # P13
    osa(f'''set s to make new slide
tell s
{TITLE("低估的红利品种，当前可以投资吗？")}
{T("• 红利星级整体在4点几星级，也在低估。",200,330,1520,80,size=30)}
{T("• 一些低估的红利指数，可以考虑定投，逐渐积累股票资产；同时做好长期定投计划，确保回到5星级时仍有资金加仓。",200,450,1520,140,size=30)}
{T("• 如果有一大笔资金，一次性大比例投入，更建议耐心等待5星级。",200,650,1520,80,size=30)}
end tell'''); print("P13 ok")
    # P14
    osa(add_image_slide('0705_p33-33.jpg')); print("P14 ok")
    # P15 图+气泡
    osa(add_image_slide('0705_p34-34.jpg'))
    osa(f'tell slide -1\n{T("参考螺丝钉红利星级及估值表，判断便宜还是贵。",560,905,820,80,size=26,rgb=(241,118,64))}\nend tell'); print("P15 ok")
    # P16 图+改文字（白图遮盖+深色新文字）
    osa(add_image_slide('0705_p21-21.jpg'))
    osa(f'''tell slide -1
make new image with properties {{file:POSIX file "{WHITE}", position:{{95, 298}}, width:1760, height:100}}
{T("所以历史上红利指数，高估比较少见，主要出现在2007年、2015年大牛市。其他时间里，多为低估和正常估值。",110,312,1720,90,size=26)}
end tell'''); print("P16 ok")
    # P17
    osa(add_image_slide('0705_p36-36.jpg')); print("P17 ok")
    # P18
    osa(f'''set s to make new slide
tell s
{TITLE("某红利指数基金的历次分红")}
{T("红利指数基金的基金分红，来自于股息分红。股息分红来自上市公司盈利的一部分，不受牛熊市涨跌影响。",470,890,1000,110,size=30)}
{SRC("数据来源：Wind（截至"+LATEST+"）",1430,1030,480)}
end tell
{add_chart(["某红利指数基金的表现（累计涨幅%）"],[p5],p5,110,300,1700,500)}'''); print("P18 ok")
    # P19
    osa(f'''set s to make new slide
tell s
{TITLE("总结")}
{T("近期红利品种表现较好，主要与市场风格轮动有关。拉长时间看，成长与价值风格的长期表现相差不大。",160,300,1600,100,size=28)}
{T("投资红利品种，适合低估时买入、持股收息。",160,420,1600,70,size=28)}
{T("判断估值时，可以股息率为主，结合盈利收益率、市盈率和市净率综合判断，也可参考螺丝钉红利星级及估值表。",160,520,1600,100,size=28)}
{T("红利指数基金既可以根据估值止盈，也可以长期持有，通过分红获得现金流。",160,650,1600,90,size=28)}
end tell'''); print("P19 ok")

    # 此时: slide1=封面, 2..n0=模板示例页, n0+1..n0+18=我的18内容页
    total = int(osa_app('return (count of slides of myDoc) as string'))
    print(f"加完18内容页, 总页数={total} (应为 {n0}+18={n0+18})")
    if total != n0 + 18:
        print("⚠️ 页数不符，为安全不删除、中止。"); sys.exit(1)
    # 删除模板示例页 slide 2..n0 (保留封面slide1 + 我的18页)
    osa(f'repeat {n0-1} times\ndelete slide 2\nend repeat')
    nfin = int(osa_app('return (count of slides of myDoc) as string'))
    print("最终页数:", nfin)
    if nfin != 19:
        print(f"⚠️ 应为19，实际{nfin}，中止。"); sys.exit(1)
    osa_app('save myDoc')
    print("已保存。DONE")

if __name__ == '__main__':
    build_all()
