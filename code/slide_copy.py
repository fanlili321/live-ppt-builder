# -*- coding: utf-8 -*-
"""跨文件复制整页：把 src_prs 的第 idx 页复制进 dst_prs(用blank_layout)。
处理形状/图片/图表的 part 关系(rels)重建。"""
import copy
from pptx.oxml.ns import qn

def copy_slide(src_prs, idx, dst_prs, blank_layout):
    src = src_prs.slides[idx]
    dst = dst_prs.slides.add_slide(blank_layout)
    # 清掉blank layout自带的占位符
    for sh in list(dst.shapes):
        sh._element.getparent().remove(sh._element)
    # 逐个复制shape element,并重建其引用的part(图片/图表)
    for sh in src.shapes:
        el = copy.deepcopy(sh._element)
        dst.shapes._spTree.append(el)
    # 源尺寸≠目标尺寸时,按比例缩放坐标+字号
    s = dst_prs.slide_width / src_prs.slide_width
    if abs(s-1) > 0.001:
        _scale(dst.shapes._spTree, s)
    # 重建 rels: 找出el里所有 r:embed / r:id 引用,复制对应part
    _relink(src, dst)
    return dst

def _scale(spTree, s):
    A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
    for off in spTree.iter(A+'off'):
        if off.get('x') is not None: off.set('x',str(int(int(off.get('x'))*s)))
        if off.get('y') is not None: off.set('y',str(int(int(off.get('y'))*s)))
    for ext in spTree.iter(A+'ext'):
        if ext.get('cx') is not None: ext.set('cx',str(int(int(ext.get('cx'))*s)))
        if ext.get('cy') is not None: ext.set('cy',str(int(int(ext.get('cy'))*s)))
    for el in spTree.iter():
        sz=el.get('sz')
        if sz and el.tag.endswith('}rPr') or (sz and el.tag.endswith('}defRPr')) or (sz and el.tag.endswith('}endParaRPr')):
            el.set('sz',str(int(int(sz)*s)))

def _relink(src, dst):
    from io import BytesIO
    dst_part = dst.part; src_part = src.part
    spTree = dst.shapes._spTree
    NS='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
    used = {}
    for el in spTree.iter():
        for attr in (NS+'embed', NS+'id', NS+'link'):
            rId = el.get(attr)
            if not rId: continue
            if rId in used: el.set(attr, used[rId]); continue
            try: src_rel = src_part.rels[rId]
            except KeyError: continue
            if src_rel.is_external:
                new_rId = dst_part.rels.get_or_add_ext_rel(src_rel.reltype, src_rel.target_ref)
            elif 'image' in src_rel.reltype:
                _, new_rId = dst_part.get_or_add_image_part(BytesIO(src_rel.target_part.blob))
            else:
                new_rId = dst_part.relate_to(src_rel.target_part, src_rel.reltype)
            used[rId] = new_rId
            el.set(attr, new_rId)

if __name__=='__main__':
    import sys
    from pptx import Presentation
    src=Presentation('/tmp/429.pptx')
    # 新建空deck(用模板)
    import ppt_engine as E
    dst=Presentation(E.TMPL_PATH)
    body=E.get_layout(dst,1,'自定义版式')
    copy_slide(src,12,dst,body)  # 429 P13
    copy_slide(src,16,dst,body)  # 429 P17(估值表页,含图片)
    for _ in range(8): E.remove_slide(dst,0)
    dst.save('/tmp/copytest.pptx')
    print('saved',len(dst.slides),'slides')
