"""导出页图后自查：▶ 中心与首行文字中心差（diff 应为 +8）、▶ 与文字间距 gap。用法: python3 check_triangles.py 导出页图目录"""
import sys, glob, numpy as np
from PIL import Image
from scipy import ndimage
for p in sorted(glob.glob(sys.argv[1]+'/*.png')):
    a=np.asarray(Image.open(p).convert('RGB')).astype(int)
    r,g,b=a[...,0],a[...,1],a[...,2]
    m=(r>200)&(g>90)&(g<160)&(b<90)
    lab,n=ndimage.label(m)
    for i,sl in enumerate(ndimage.find_objects(lab)):
        h=sl[0].stop-sl[0].start; w=sl[1].stop-sl[1].start
        if not(30<=h<=50 and 25<=w<=50): continue
        y0,y1,x0,x1=sl[0].start,sl[0].stop,sl[1].start,sl[1].stop
        if x0>400: continue
        tc=(y0+y1)/2
        # first text line to the right
        reg=a[y0-40:y1+40, x1+5:x1+700]; dark=reg.sum(2)<300
        rows=np.where(dark.sum(1)>3)[0]
        # band containing rows nearest tc
        if len(rows)==0: continue
        rows=rows+y0-40
        bands=[]; s=rows[0]; prev=rows[0]
        for y in rows[1:]:
            if y>prev+3: bands.append((s,prev)); s=y
            prev=y
        bands.append((s,prev))
        bd=min(bands,key=lambda bb:abs((bb[0]+bb[1])/2-tc))
        cols=np.where((a[bd[0]:bd[1]+1, x1:x1+300].sum(2)<300).any(0))[0]
        gx=x1+cols[0] if len(cols) else -1
        print(p.split('/')[-1], f'tri x{x0}-{x1} y{y0}-{y1} c{tc:.0f} | line y{bd[0]}-{bd[1]} c{(bd[0]+bd[1])/2:.0f} | diff {(bd[0]+bd[1])/2-tc:+.0f} | gap {gx-x1}')
