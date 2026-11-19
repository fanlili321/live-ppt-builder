# 版式小工具：PNG 图形层 + 生成 Keynote AppleScript 文字/图片指令（487/488 期）
# 用法：新页的底色块、圆角卡片、序号圆、分隔线画成一张透明 PNG（Layer，按幻灯片 1920x1080 坐标），
#      文字一律做成 Keynote 原生文本框（SlideScript.text，思源字体，宽度贴合文字），保证可编辑。
#   from kit import *
#   sl = SlideScript(页序号, '标题里的一段字')      # 页序号=slide 序号（含隐藏页），标题不符即报错
#   L = Layer('p7_steps'); L.circle(92, 415, 56, ORANGE); sl.image(L)
#   sl.text('1', 92, 415, size=36, color=WHITE); sl.text('扫码进入…', 132, 415, size=36, align='left')
#   build_script('/绝对路径/成品.key', [sl], 'tmp/x.applescript', count=总页数)   # 再 osascript 执行（文件需已在 Keynote 打开）
# 例：488 期 P5 五张信号卡片、P7/P12 序号步骤（复用 P18 格式）。改前先留备份；改坏了用备份整份还原再重跑。
import os
from PIL import Image, ImageDraw, ImageFont

# 生成的 PNG 放这里：环境变量 KIT_ASSET 指定，默认「当前目录/assets」（放在 output/NNN_claude输出/ 下运行）
ASSET = os.environ.get('KIT_ASSET') or os.path.join(os.getcwd(), 'assets')
os.makedirs(ASSET, exist_ok=True)
FONTDIR = os.path.expanduser('~/Library/Fonts')
FONTS = {'Bold': 'SourceHanSansCN-Bold', 'Normal': 'SourceHanSansCN-Normal',
         'Medium': 'SourceHanSansCN-Medium', 'Regular': 'SourceHanSansCN-Regular',
         'Heavy': 'SourceHanSansCN-Heavy'}

# 配色（取自本 PPT 原页）
ORANGE = '#E3792F'; YELLOW = '#F2B13D'; BLUE = '#6A75A3'
ORANGE_L = '#FDF0E8'; BLUE_L = '#E8EEFC'; BUBBLE = '#FFDDC4'; BUBBLE_S = '#EB873F'
ORANGE_S = '#FBC299'; BLUE_S = '#BFC8E8'; FRAME = '#FEAD00'
BLACK = '#000000'; DARK = '#333333'; GREY = '#4B4B4B'; LGREY = '#9A9A9A'; WHITE = '#FFFFFF'
LINE_300 = '#EA7C2D'; LINE_800 = '#F9DB57'; LINE_A500 = '#929292'; LINE_ALL = '#EA4026'

S = 3  # PNG 超采样倍数


def rgb(h):
    h = h.lstrip('#'); return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def c16(h):
    return '{%d, %d, %d}' % tuple(v * 257 for v in rgb(h))


def tw(text, size, weight='Bold'):
    f = ImageFont.truetype(os.path.join(FONTDIR, FONTS[weight] + '.otf'), int(size * 4))
    return max(f.getlength(line) for line in text.split('\n')) / 4


class Layer:
    """一组图形 → 一张透明 PNG（坐标为幻灯片 1920x1080 坐标）"""

    def __init__(self, name):
        self.name = name; self.ops = []

    def rrect(self, x, y, w, h, r, fill=None, stroke=None, sw=0, corners=None):
        self.ops.append(('rrect', x, y, w, h, r, fill, stroke, sw, corners)); return self

    def circle(self, cx, cy, d, fill):
        self.ops.append(('ellipse', cx - d / 2, cy - d / 2, d, d, fill)); return self

    def poly(self, pts, fill):
        self.ops.append(('poly', pts, fill)); return self

    def line(self, x1, y1, x2, y2, color, width):
        self.ops.append(('line', x1, y1, x2, y2, color, width)); return self

    def bbox(self):
        xs, ys = [], []
        for o in self.ops:
            if o[0] in ('rrect', 'ellipse'):
                pad = (o[8] if o[0] == 'rrect' else 0) / 2 + 2
                xs += [o[1] - pad, o[1] + o[3] + pad]; ys += [o[2] - pad, o[2] + o[4] + pad]
            elif o[0] == 'poly':
                xs += [p[0] for p in o[1]]; ys += [p[1] for p in o[1]]
            elif o[0] == 'line':
                xs += [o[1] - o[6], o[3] + o[6]]; ys += [o[2] - o[6], o[4] + o[6]]
        return int(min(xs)) - 2, int(min(ys)) - 2, int(max(xs)) + 3, int(max(ys)) + 3

    def render(self):
        x0, y0, x1, y1 = self.bbox(); W, H = x1 - x0, y1 - y0
        im = Image.new('RGBA', (W * S, H * S), (0, 0, 0, 0)); d = ImageDraw.Draw(im)
        T = lambda x, y: ((x - x0) * S, (y - y0) * S)
        for o in self.ops:
            if o[0] == 'rrect':
                _, x, y, w, h, r, fill, stroke, sw, corners = o
                a = T(x, y); b = T(x + w, y + h)
                kw = dict(radius=r * S, fill=rgb(fill) if fill else None)
                if stroke: kw.update(outline=rgb(stroke), width=int(sw * S))
                if corners: kw['corners'] = corners
                d.rounded_rectangle([a, (b[0] - 1, b[1] - 1)], **kw)
            elif o[0] == 'ellipse':
                _, x, y, w, h, fill = o
                d.ellipse([T(x, y), T(x + w, y + h)], fill=rgb(fill))
            elif o[0] == 'poly':
                d.polygon([T(*p) for p in o[1]], fill=rgb(o[2]))
            elif o[0] == 'line':
                _, a1, b1, a2, b2, col, wd = o
                d.line([T(a1, b1), T(a2, b2)], fill=rgb(col), width=int(wd * S))
        path = os.path.join(ASSET, self.name + '.png'); im.save(path)
        return path, (x0, y0, W, H)


class SlideScript:
    def __init__(self, idx, expect_title):
        self.idx = idx; self.title = expect_title; self.lines = []

    def raw(self, s):
        self.lines.append(s)

    def image(self, layer):
        path, (x, y, w, h) = layer.render()
        self.lines.append(
            f'tell s to make new image with properties {{file:(POSIX file "{path}") as alias, position:{{{x}, {y}}}, width:{w}, height:{h}}}')

    def text(self, txt, x, cy, w=None, size=40, weight='Bold', color=BLACK, align='center', pad=24):
        """align=center: 以 x 为中心; left: 以 x 为左边界（文字框宽度贴合文字，新建框默认居中对齐）"""
        width = w or (tw(txt, size, weight) + pad)
        if align == 'left':
            left = x - 10
        elif align == 'right':
            left = x - width + 10
        else:
            left = x - width / 2
        parts = txt.split('\n')
        astxt = ' & return & '.join('"%s"' % p for p in parts)
        self.lines += [
            f'tell s to set t to make new text item with properties {{object text:({astxt}), position:{{{int(left)}, 1500}}, width:{int(width)}, height:100}}',
            f'set font of object text of t to "{FONTS[weight]}"',
            f'set size of object text of t to {size}',
            f'set color of object text of t to {c16(color)}',
            f'set width of t to {int(width)}',
            f'set h to height of t',
            f'set position of t to {{{int(left)}, ({cy} - h / 2) as integer}}',
        ]

    def emit(self):
        head = [f'set s to slide {self.idx} of d',
                f'if (object text of default title item of s) does not contain "{self.title}" then error "slide {self.idx} title mismatch"']
        return '\n'.join(head + self.lines)


def build_script(doc_path, slides, out, count=37):
    body = '\n\n'.join(sl.emit() for sl in slides)
    src = f'''with timeout of 900 seconds
tell application "Keynote"
set d to missing value
repeat with dd in documents
 if (POSIX path of (file of dd as alias)) is "{doc_path}" then set d to document id (id of dd)
end repeat
if d is missing value then error "doc not found"
if (count of slides of d) is not {count} then error "slide count"
{body}

save d
end tell
end timeout
return "ok"
'''
    open(out, 'w').write(src)
