"""把导出的幻灯片图片里的走势图区域裁出来做“毛玻璃”模糊（图例和起止日期保持清晰）。

用法：python3 blur_chart.py 幻灯片.png 输出.png 裁剪顶边Y [左X 宽 高]
默认裁 (122, top, 122+922, top+537)，即第26期 p6/p7 的模糊图位置。
模糊区：左侧Y轴标签 (0,110)-(170,475)；折线区 (0,160)-(全宽,470)；图例第三行右侧 (470,110)-(全宽,160)。
图表布局不同的页，先看一眼裁出来的图再调这几个框。
"""
import sys
from PIL import Image, ImageFilter, ImageDraw

src, out, top = sys.argv[1], sys.argv[2], int(sys.argv[3])
left = int(sys.argv[4]) if len(sys.argv) > 4 else 122
w = int(sys.argv[5]) if len(sys.argv) > 5 else 922
h = int(sys.argv[6]) if len(sys.argv) > 6 else 537

im = Image.open(src).convert('RGB')
crop = im.crop((left, top, left + w, top + h))


def blur(box):
    r = crop.crop(box).filter(ImageFilter.GaussianBlur(10))
    crop.paste(r, box[:2])


for b in [(0, 110, 170, 475), (0, 160, crop.width, 470), (470, 110, crop.width, 160)]:
    blur(b)
ImageDraw.Draw(crop).rounded_rectangle((2, 2, crop.width - 3, crop.height - 3), radius=8,
                                       outline=(245, 198, 122), width=4)
crop.save(out)
print(out, crop.size)
