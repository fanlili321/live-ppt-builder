"""导出页图后自查：气泡四边留白（L/R/T/B，差应≤3px）。用法: python3 check_bubble_margins.py 导出页图目录"""
import sys, glob, numpy as np
from PIL import Image
from scipy import ndimage
for p in sorted(glob.glob(sys.argv[1]+'/*.png')):
    a=np.asarray(Image.open(p).convert('RGB')).astype(int)
    m=(abs(a[...,0]-255)<6)&(abs(a[...,1]-221)<10)&(abs(a[...,2]-196)<12)
    lab,n=ndimage.label(ndimage.binary_closing(m,iterations=8))
    for i,sl in enumerate(ndimage.find_objects(lab)):
        y0,y1,x0,x1=sl[0].start,sl[0].stop,sl[1].start,sl[1].stop
        if (x1-x0)*(y1-y0)<60000: continue
        # body: rows where fill width is large -> use column histogram of fill to exclude pointer
        reg=(lab[sl]==i+1)
        cols=reg.sum(0); body=np.where(cols>0.8*cols.max())[0]; bx0,bx1=x0+body[0],x0+body[-1]
        rows=reg.sum(1); br=np.where(rows>0.8*rows.max())[0]; by0,by1=y0+br[0],y0+br[-1]
        sub=a[by0:by1,bx0:bx1]; dark=(sub.sum(2)<250)
        ys,xs=np.where(dark)
        if len(xs)==0: continue
        L=xs.min(); R=(bx1-bx0)-xs.max(); T=ys.min(); B=(by1-by0)-ys.max()
        print(p.split('/')[-1], f'body {bx0},{by0}-{bx1},{by1}  L{L} R{R} T{T} B{B}  dx{(R-L)//2} dy{(B-T)//2}')
