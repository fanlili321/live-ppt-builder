# 估值表截图「突出展示某个指数」固定格式（490 期定，2026-11-26 用户批注）：
#   整张估值表（右侧基金代码列打码）+ 表头放大条紧贴在该指数放大条正上方（两条连成一块）+ 表上该指数行红框 + 红虚线连到放大块；左侧标题 + 浅橙气泡。
#   kit 图层（表格、放大条带 #FEAD00 6pt 边烤进图层）+ 原生气泡文字。用法见文件末尾 page(...) 调用：行号=该指数行在原图的 y 范围，strip_y=放大块顶部 y。
#   新图层插入后会盖住旧文字：界面里选中图层 → 排列 ▸ 移到最后面。
import sys
sys.path.insert(0,'/Users/fanlili/.claude/skills/live-ppt-builder/scripts')
from kit import *
from PIL import Image
DOC=sys.argv[1]
WITH_TEXT='--text' in sys.argv
def prep(src,name,row,hdr=(39,66)):
    im=Image.open(src).convert('RGB'); W,H=im.size
    # mosaic fund code columns
    reg=im.crop((400,hdr[1],W,H)); sm=reg.resize((max(1,reg.width//10),max(1,reg.height//10)),Image.BILINEAR).resize(reg.size,Image.NEAREST)
    im.paste(sm,(400,hdr[1]))
    p=f'assets/{name}_table.png'; im.save(p)
    h=im.crop((0,hdr[0],W,hdr[1])); ph=f'assets/{name}_hdr.png'; h.save(ph)
    r=im.crop((0,row[0],W,row[1])); pr=f'assets/{name}_row.png'; r.save(pr)
    return p,ph,pr,W,H
def page(idx,title,src,name,row,strip_y,bubble,bw=800):
    p,ph,pr,W,H=prep(src,name,row)
    th=960; tw_=W*th/H; tx=1180; ty=70; k=th/H
    sl=SlideScript(idx,title); L=Layer(name+'_lay')
    L.img(p,tx,ty,tw_,th)
    sw=760; sx=tx+tw_/2-sw/2
    W0=W
    hh=(66-39)*sw/W0+6; rh=(row[1]-row[0])*sw/W0+6
    # 表头放大条紧贴在指数放大条上面一行（2026-11-26 用户定）
    sy_h=strip_y; sy_r=strip_y+hh
    ry0=ty+row[0]*k; ry1=ty+row[1]*k
    L.rrect(tx-2,ry0-2,tw_+4,ry1-ry0+4,0,stroke='#E0402A',sw=3)
    for a,b in [((tx,ry0),(sx,sy_h)),((tx+tw_,ry0),(sx+sw,sy_h)),((tx,ry1),(sx,sy_r+rh)),((tx+tw_,ry1),(sx+sw,sy_r+rh))]:
        n=14
        for i in range(0,n,2):
            L.line(a[0]+(b[0]-a[0])*i/n,a[1]+(b[1]-a[1])*i/n,a[0]+(b[0]-a[0])*(i+1)/n,a[1]+(b[1]-a[1])*(i+1)/n,'#E0402A',2.5)
    L.rrect(sx-3,sy_h-3,sw+6,hh+rh+6,0,fill='#E0402A')
    L.img(ph,sx,sy_h,sw,hh,border=0)
    L.img(pr,sx,sy_r,sw,rh,border=0)
    # bubble
    bx,by,bh=110,560,190
    L.rrect(bx,by,bw,bh,24,fill=BUBBLE,stroke=BUBBLE_S,sw=2)
    L.poly([(bx+bw-2,by+bh/2-26),(bx+bw+60,by+bh/2),(bx+bw-2,by+bh/2+26)],BUBBLE)
    L.line(bx+bw-1,by+bh/2-26,bx+bw+60,by+bh/2,BUBBLE_S,2); L.line(bx+bw+60,by+bh/2,bx+bw-1,by+bh/2+26,BUBBLE_S,2)
    sl.image(L)
    if WITH_TEXT: sl.text(bubble,bx+bw/2,by+bh/2,w=bw-88,size=40,weight='Medium')
    return sl
s1=page(11,'原因一','outline/p10_val_q2.png','p11',(1340,1370),860,'2026年2季度，成长风格大涨，估值达到高估。',bw=660)
s2=page(16,'螺丝钉星级和','outline/p15_val_now.png','p16',(513,543),330,'目前创业板指回到低估。')
build_script(DOC,[s1,s2],'tmp_val.applescript',count=21)
