#!/usr/bin/env python3
"""把来源期 .key 的若干页整页复制到底稿指定位置（Keynote 界面拷贝/粘贴）。

AppleScript 不能跨文档 duplicate slide，所以走界面：
  来源副本：AppleScript 把第 p 页挪到第 1 页 → 真实鼠标点第 1 张缩略图（焦点进导航器）→ 编辑▸拷贝 → 挪回原位；
  底稿：点第 1 张缩略图 → 编辑▸粘贴（落在第 2 页）→ 核对标题/图表数/页数 → AppleScript 挪到 --after 之后。
（2026-11-10 实测：脚本设当前页会让焦点离开导航器、拷贝变灰；方向键/End/Home 偶尔丢键或页码回报滞后，所以只点第 1 张。）

前提：小写「claude」（Claude Code 命令行）有「辅助功能」权限（用户本人在系统设置里开）。
运行时 Keynote 会被切到前台、鼠标会动，期间不要操作电脑。

用法：
  python3 copy_slides.py --check
  python3 copy_slides.py --src 455期副本.key --pages 10,12 --dst work.key --after 5

规则：src/dst 都必须是副本；已打开且 modified=true 的文档拒绝操作（src 除外：本脚本自己移过页）。
每粘一页都核对：底稿页数 +1、新页标题 = 来源页标题、图表数一致；不一致立即停止。
来源副本最后不保存关闭。2026-11-10 在 484→485 副本上实测通过（思维导图页、双图表页）。
"""
import argparse
import json
import subprocess
import sys
from pathlib import Path

JXA = r'''
ObjC.import('CoreGraphics');
ObjC.import('AppKit');
function click(x,y){
  const p=$.CGPointMake(x,y);
  $.CGEventPost($.kCGHIDEventTap,$.CGEventCreateMouseEvent(null,$.kCGEventMouseMoved,p,0)); delay(0.1);
  $.CGEventPost($.kCGHIDEventTap,$.CGEventCreateMouseEvent(null,$.kCGEventLeftMouseDown,p,0)); delay(0.05);
  $.CGEventPost($.kCGHIDEventTap,$.CGEventCreateMouseEvent(null,$.kCGEventLeftMouseUp,p,0));
}
function pbTypes(){
  const t=$.NSPasteboard.generalPasteboard.types; const o=[];
  for(let i=0;i<t.count;i++) o.push(ObjC.unwrap(t.objectAtIndex(i)));
  return o;
}
function run(argv){
  const [win,item]=argv;
  const kn=Application('Keynote'); const d=kn.documents.byName(win);
  d.currentSlide=d.slides[0]; delay(0.6);              // 导航器滚回顶部
  kn.activate(); delay(0.8);
  const se=Application('System Events'); const pr=se.processes.byName('Keynote');
  const w=pr.windows.byName(win);
  w.actions.byName('AXRaise').perform(); delay(0.5);
  const sa=w.splitterGroups[0].scrollAreas[0];       // 幻灯片导航器
  const [px,py]=sa.position(); const [sx,sy]=sa.size();
  click(px+sx/2, py+45); delay(0.8);                  // 点第 1 张缩略图，焦点进导航器
  const cur=d.currentSlide.slideNumber();
  if(cur!==1) return JSON.stringify({ok:false, why:'点击后不在第1页', cur});
  if(!pr.frontmost()){kn.activate(); delay(0.6);}
  const mi=pr.menuBars[0].menus.byName('编辑').menuItems.byName(item);
  if(!mi.enabled()) return JSON.stringify({ok:false, why:'菜单不可用（焦点不在导航器）'});
  mi.click(); delay(1.5);
  return JSON.stringify({ok:true, cur, front:pr.frontmost(),
    slideOnClipboard: pbTypes().indexOf('com.apple.iWork.Keynote.pasteboardState.hasSlideNodes')>=0});
}
'''


def osa(script: str, timeout: int = 120) -> str:
    r = subprocess.run(['osascript', '-'], input=script, capture_output=True, text=True, timeout=timeout)
    if r.returncode != 0:
        raise RuntimeError(r.stderr.strip())
    return r.stdout.strip()


def jxa(win: str, item: str) -> dict:
    r = subprocess.run(['osascript', '-l', 'JavaScript', '-', win, item], input=JXA,
                       capture_output=True, text=True, timeout=120)
    if r.returncode != 0:
        raise RuntimeError(r.stderr.strip())
    return json.loads(r.stdout.strip().splitlines()[-1])


def q(s: str) -> str:
    return s.replace('\\', '\\\\').replace('"', '\\"')


def check_access() -> bool:
    try:
        osa('tell application "System Events" to tell process "Keynote" to get name of every window')
        return True
    except RuntimeError as e:
        print('辅助功能权限未开：', e)
        print('请用户在 系统设置 → 隐私与安全性 → 辅助功能 里打开小写的「claude」（Claude Code 命令行程序；'
              '橙色图标的 Claude App 开了不够），开完重启 Claude App。')
        return False


def open_doc(path: Path, allow_modified=False) -> str:
    return osa(f'''
tell application "Keynote"
  with timeout of 100 seconds
    set d to open POSIX file "{q(str(path))}"
    if (modified of d) and not {str(allow_modified).lower()} then error "文档有未保存修改：" & (name of d)
    return name of d
  end timeout
end tell''')


def info(doc: str, n) -> dict:
    """第 n 页（或 'last'）的标题、图表数，以及文档页数。"""
    ref = '(count of slides of d)' if n == 'last' else str(n)
    out = osa(f'''
tell application "Keynote"
  set d to document "{q(doc)}"
  set s to slide {ref} of d
  set t to ""
  try
    set t to object text of default title item of s
  end try
  if t is "" then
    repeat with x in text items of s
      set t to object text of x
      if t is not "" then exit repeat
    end repeat
  end if
  return ((count of slides of d) as text) & "|||" & ((count of charts of s) as text) & "|||" & t
end tell''')
    cnt, charts, title = out.split('|||', 2)
    return {'count': int(cnt), 'charts': int(charts), 'title': title.strip()}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--check', action='store_true')
    ap.add_argument('--src', type=Path)
    ap.add_argument('--pages', help='来源页码，逗号分隔，按顺序依次贴在 --after 之后')
    ap.add_argument('--dst', type=Path)
    ap.add_argument('--after', type=int, help='粘贴到底稿第几页之后')
    a = ap.parse_args()

    if not check_access():
        sys.exit(2)
    if a.check:
        print('辅助功能权限 OK')
        return
    src = open_doc(a.src.expanduser().resolve())
    dst = open_doc(a.dst.expanduser().resolve())
    pos = a.after
    for p in [int(x) for x in a.pages.split(',')]:
        want = info(src, p)
        mv = lambda doc, a_, how, b_: osa(f'tell application "Keynote" to move slide {a_} of document "{q(doc)}" to {how} slide {b_} of document "{q(doc)}"')
        # 来源副本：第 p 页挪到第 1 页 → 点第 1 张缩略图 → 拷贝 → 挪回原位
        if p != 1:
            mv(src, p, 'before', 1)
        r1 = jxa(src, '拷贝')
        if p != 1:
            mv(src, 1, 'after', p)
        if not (r1.get('ok') and r1.get('slideOnClipboard')):
            sys.exit(f'!! 来源P{p} 拷贝失败：{r1}')
        # 底稿：选中第 1 页 → 粘贴（落在第 2 页）→ 挪到 pos 之后
        n0 = info(dst, 1)['count']
        r2 = jxa(dst, '粘贴')
        if not r2.get('ok'):
            sys.exit(f'!! 粘贴失败：{r2}')
        new2 = info(dst, 2)
        if not (new2['count'] == n0 + 1 and new2['title'] == want['title'] and new2['charts'] == want['charts']):
            sys.exit(f'!! 粘贴核对不一致，停止（底稿未保存）：来源P{p} {want} → 底稿P2 {new2}')
        if pos != 1:
            mv(dst, 2, 'after', pos + 1)
        got = info(dst, pos + 1)
        if got['title'] != want['title']:
            sys.exit(f'!! 移动后位置不对（底稿未保存）：底稿P{pos + 1} {got}')
        print(f'OK 来源P{p}「{want["title"][:24]}」(图表{want["charts"]}) → 底稿P{pos + 1}，底稿共 {got["count"]} 页')
        pos += 1
    osa(f'tell application "Keynote" to save document "{q(dst)}"')
    osa(f'tell application "Keynote" to close document "{q(src)}" saving no')
    print('底稿已保存；来源副本已关闭（未保存）。')


if __name__ == '__main__':
    main()
