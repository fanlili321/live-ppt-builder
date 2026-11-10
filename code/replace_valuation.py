# -*- coding: utf-8 -*-
"""复用「星级&指数估值表」页：只替换右侧估值表图(可选+星级徽标)，其余不动。
用法: python3 replace_valuation.py <估值表图.png> [星级徽标图.png] [-o 输出.pptx]
默认基于模板 templates/456人工成品_模板.pptx 的第16页。"""
import argparse
import os
import sys
from pptx import Presentation

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE = os.path.join(BASE, 'templates', '456人工成品_模板.pptx')
VAL_SLIDE = 15          # 第16页(0基)
VAL_PIC_IDX = 0         # 估值表大图
BADGE_PIC_IDX = 9       # 星级徽标(可选)

def replace_pic_blob(slide, shape_idx, img_path):
    pic = slide.shapes[shape_idx]
    rId = pic._element.blip_rId
    img_part = pic.part.related_part(rId)
    with open(img_path, 'rb') as f:
        img_part._blob = f.read()

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('valuation_img')
    ap.add_argument('badge_img', nargs='?')
    ap.add_argument('-o', '--out', default=os.path.expanduser('~/Downloads/估值表页_更新.pptx'))
    ap.add_argument('-t', '--template', default=TEMPLATE)
    a = ap.parse_args()
    prs = Presentation(a.template)
    sl = prs.slides[VAL_SLIDE]
    replace_pic_blob(sl, VAL_PIC_IDX, a.valuation_img)
    if a.badge_img:
        replace_pic_blob(sl, BADGE_PIC_IDX, a.badge_img)
    prs.save(a.out)
    print(f'✅ 估值表已替换 → {a.out}')

if __name__ == '__main__':
    main()
