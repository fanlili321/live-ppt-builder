"""在 iCloud 协作在线文档里整页替换/插入（486 期 2026-11-13 实测可用）。
⚠️ 只在用户明确同意替换在线文档时用；定稿后在线文档一律不动（安全铁律第 15 条）。
前提：源文件与在线文档都已在 Keynote 打开；源文件放在 output/NNN_claude输出/tmp/（不能放 /tmp），在线文档当前选中第 1 页（粘贴落在第 2 页）。
每一步都校验：粘贴后等 3 秒（协作文档页数更新慢）、页数 +1、页面签名（标题 + 文本框/图片/形状数 + 全部文字的哈希）与源页一致；
移到目标位置后再校验前后页签名，最后删除旧页并核对页数。任一校验失败立即 AssertionError 停止，不继续。
用法: python3 replace_slides_online.py --src 源.key文档名 --dst 在线文档名 r:源页:在线旧页 i:源页:插在在线第N页之后 …
  例: python3 replace_slides_online.py --src src_for_online.key --dst '标题.key' r:11:11 r:13:13 i:14:13
  顺序执行；插入会让后面的页号 +1，操作列表按执行后的页号写。替换前先 cp -R 留在线文档快照，替换后导出页图与源文件逐页像素比对。"""
import sys, time, hashlib, os, argparse
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import copy_slides as c
ap = argparse.ArgumentParser(); ap.add_argument('--src', required=True); ap.add_argument('--dst', required=True); ap.add_argument('ops', nargs='+')
A = ap.parse_args(); SRC, DST = A.src, A.dst
def sig(doc, n):
    o = c.osa(f'''tell application "Keynote"
set s to slide {n} of document "{DST if doc == 'd' else SRC}"
set t to ""
try
set t to object text of default title item of s
end try
set AppleScript's text item delimiters to "¦"
set a to (object text of every text item of s) as text
return t & "|" & (count of text items of s) & "|" & (count of images of s) & "|" & (count of shapes of s) & "|" & a
end tell''')
    return o.split('|')[0] + '#' + hashlib.md5(o.encode()).hexdigest()[:8]
def count(): return c.info(DST, 1)['count']
def paste_new(p):
    want = sig('s', p); n0 = count()
    c.osa(f'tell application "Keynote" to move slide {p} of document "{SRC}" to before slide 1 of document "{SRC}"')
    r1 = c.jxa(SRC, '拷贝')
    c.osa(f'tell application "Keynote" to move slide 1 of document "{SRC}" to after slide {p} of document "{SRC}"')
    assert r1.get('slideOnClipboard') and sig('s', p) == want, r1
    c.jxa(DST, '粘贴'); time.sleep(3)
    assert count() == n0 + 1 and sig('d', 2) == want, ('粘贴不对', count(), sig('d', 2), want)
    return want, n0
def replace(p, old):
    before = sig('d', old); want, n0 = paste_new(p)
    assert before != want, '新旧页无法区分'
    c.osa(f'tell application "Keynote" to move slide 2 of document "{DST}" to after slide {old + 1} of document "{DST}"')
    assert sig('d', old) == before and sig('d', old + 1) == want, ('位置不对', sig('d', old), sig('d', old + 1))
    c.osa(f'tell application "Keynote" to delete slide {old} of document "{DST}"')
    assert count() == n0 and sig('d', old) == want
    print('替换', p, '→', old, want, flush=True)
def insert(p, after):
    anchor = sig('d', after); want, n0 = paste_new(p)
    c.osa(f'tell application "Keynote" to move slide 2 of document "{DST}" to after slide {after + 1} of document "{DST}"')
    assert sig('d', after) == anchor and sig('d', after + 1) == want, '位置不对'
    print('插入', p, '→', after + 1, want, flush=True)
for op in A.ops:
    k, x, y = op.split(':'); (replace if k == 'r' else insert)(int(x), int(y))
