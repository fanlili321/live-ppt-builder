#!/usr/bin/env python3
"""定稿后清理一期的中间文件：移到废纸篓（可在访达里「放回原处」），不永久删除。

保留：定稿 .key（--keep 指定，可多个）、过程记录 md、scripts/、data/ 里的原始数据和计算结果（json/csv/xlsx/txt 等小文件）、data/received/。
移到废纸篓：
  - data/改前备份_*.key（所有改前备份）
  - tmp/、src/（历史期 .key 复制件，缓存 history_keys/ 里还有）
  - export*/、review/ 下的页图和副本、pages*/、outline/img/
  - 除 --keep 以外的 *.key（v1、v2、快照等中间版本）
不碰：iCloud 在线文档、history_keys/ 缓存、用户自己的文件、Keynote 正开着的文件（检测到就跳过并提示）。

用法：
  python3 cleanup_episode.py output/NNN_claude输出 --keep 定稿.key            # 只列清单和大小（默认 dry-run）
  python3 cleanup_episode.py output/NNN_claude输出 --keep 定稿.key --apply    # 执行：移到废纸篓
"""
import argparse, os, shutil, subprocess, time
from pathlib import Path

def size(p: Path) -> int:
    if p.is_file(): return p.stat().st_size
    return sum(f.stat().st_size for f in p.rglob('*') if f.is_file())

def human(n: int) -> str:
    for u in ('B', 'K', 'M', 'G'):
        if n < 1024: return f'{n:.0f}{u}'
        n /= 1024
    return f'{n:.1f}T'

def open_in_keynote():
    try:
        r = subprocess.run(['osascript', '-e', 'tell application "System Events" to set k to (exists process "Keynote")',
                            '-e', 'if k then tell application "Keynote" to get POSIX path of (file of every document as list)'],
                           capture_output=True, text=True, timeout=30)
        return {Path(x.strip().rstrip('/')).resolve() for x in r.stdout.split(',') if x.strip()}
    except Exception:
        return set()

def main():
    ap = argparse.ArgumentParser(); ap.add_argument('folder'); ap.add_argument('--keep', action='append', default=[])
    ap.add_argument('--apply', action='store_true'); a = ap.parse_args()
    root = Path(a.folder).resolve(); keep = {(root / k).resolve() for k in a.keep}
    if not keep or not all(k.exists() for k in keep):
        raise SystemExit('必须用 --keep 指定存在的定稿文件（相对该期文件夹）')
    targets = []
    for p in root.iterdir():
        n = p.name
        if p.resolve() in keep: continue
        if n in ('tmp', 'src') or n.startswith('export') or n.startswith('pages') or n == 'review':
            targets.append(p)
        elif p.suffix == '.key':
            targets.append(p)
    d = root / 'data'
    if d.exists():
        targets += [p for p in d.iterdir() if p.name.startswith('改前备份') or p.suffix == '.key']
    oi = root / 'outline' / 'img'
    if oi.exists(): targets.append(oi)
    opened = open_in_keynote()
    total = 0
    for t in sorted(targets):
        if t.resolve() in opened:
            print('跳过（Keynote 正开着）', t.relative_to(root)); continue
        s = size(t); total += s
        print(f'{human(s):>6}  {t.relative_to(root)}')
    print(f'合计 {human(total)}；保留：' + '、'.join(str(k.relative_to(root)) for k in keep))
    if not a.apply:
        print('（dry-run，未移动；确认后加 --apply）'); return
    trash = Path.home() / '.Trash'
    for t in targets:
        if t.resolve() in opened or not t.exists(): continue
        dst = trash / t.name
        if dst.exists(): dst = trash / f'{t.stem}_{int(time.time())}{t.suffix}'
        shutil.move(str(t), str(dst))
    print(f'已移到废纸篓，释放约 {human(total)}（可在访达废纸篓里「放回原处」，未清空废纸篓）')

if __name__ == '__main__':
    main()
