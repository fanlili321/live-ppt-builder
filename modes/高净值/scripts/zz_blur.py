"""运营周报「指数增强xx期」截图 → 高净值页用的模糊图（第27期第22页做法）。
用法：python3 zz_blur.py 截图.jpg 输出.png
裁掉底部「成立以来跑赢业绩基准」淡紫色条，整图高斯模糊（图例含产品名，一起糊），
首尾日期按像素贴回清晰，外框 #FEAD00（PNG 1800×1120，页面放 868×540 时约 6 点）。"""
import sys
import numpy as np
from PIL import Image, ImageDraw, ImageFilter

src, out = sys.argv[1], sys.argv[2]
im = Image.open(src).convert('RGB'); a = np.array(im).astype(int)
lav = ((abs(a[..., 0]-228) < 14) & (abs(a[..., 2]-246) < 12) & (a[..., 2]-a[..., 1] > 8)).sum(1)
top = next(y for y in range(len(lav)) if lav[y] > im.width*0.5)
g = (a.max(2) < 190) & (abs(a[..., 0]-a[..., 2]) < 20)
ys = [y for y in range(top-50, top) if g[y].sum() > 5]
y0, y1 = min(ys), max(ys)
xs = np.where(g[y0:y1+1].any(0))[0]
cl = []; s = p = xs[0]
for x in xs[1:]:
    if x-p > 40: cl.append((s, p)); s = x
    p = x
cl.append((s, p))
c = im.crop((0, 0, im.width, top-12)); bg = tuple(int(v) for v in a[top-20, 5])
W, H = 1800, 1120
k = min((W-60)/c.width, (H-60)/c.height)
cs = c.resize((round(c.width*k), round(c.height*k)), Image.LANCZOS)
ox, oy = (W-cs.width)//2, (H-cs.height)//2
sharp = Image.new('RGB', (W, H), bg); sharp.paste(cs, (ox, oy))
cv = sharp.filter(ImageFilter.GaussianBlur(9))
for x0, x1 in cl:
    box = (int(ox+(x0-6)*k), int(oy+(y0-5)*k), int(ox+(x1+6)*k), int(oy+(y1+5)*k))
    cv.paste(sharp.crop(box), box[:2])
ImageDraw.Draw(cv).rectangle([0, 0, W-1, H-1], outline=(254, 173, 0), width=12)
cv.save(out); print(out, '日期框', cl)
