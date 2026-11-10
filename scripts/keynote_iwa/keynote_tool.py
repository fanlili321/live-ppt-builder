#!/usr/bin/env python3
"""Keynote (.key) 直改工具：读/改原生图表数据、定位页面、安全写回、核验图片引用。

为什么需要它：Keynote 的 AppleScript 改不了已有图表的数据（-10006）。本工具直接改 .key
包内 Index/Slide-*.iwa 里的图表数据网格，图表样式（配色/坐标轴/图例/标注）100% 保留。

铁律：
- 只改副本（--out 必须是新文件），改完用 Keynote 打开导出核对，再替换正式文件。
- 写回只用 zip 命令行逐条替换条目，绝不用 Python zipfile 整包重打——
  Keynote 的中文文件名（如「已粘贴的图像」「图层 1」）不带 UTF-8 标记，
  Python zipfile 重打会把文件名写成乱码，Keynote 找不到图片 → 页面空白，再保存就永久丢图。

子命令：
  find      <key> <文字>                         哪个 Slide-*.iwa 含这段文字
  charts    <key> [--slide-file F]               列出图表：系列名、行数、首末行
  get       <key> --slide-file F [--grid N] [--out x.csv]   导出图表数据为 CSV
  set       <key> --slide-file F [--grid N] --csv x.csv --out new.key
            CSV 第一行=表头（第一格空/任意，其余=系列名，须与原图一致），
            之后每行=行名（日期）, 各系列数值（空=空单元格）。行数可以和原图不同。
  text      <key> --slide-file F --old 文字 --new 文字 --out new.key   （仅限等长UTF-8替换）
  verify    <before.key> <after.key>             条目增删改 + 图片引用完整性
  restore   <key> --from <backup.key> --out new.key   从备份补回被引用但缺失的 Data 文件
"""
import argparse, csv, os, re, shutil, struct, subprocess, sys, tempfile, uuid, zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from iwa import read_iwa, write_iwa  # noqa: E402
from pb import rv, ev, fields, objects  # noqa: E402


# ---------- zip helpers (name-safe) ----------
def entry_name(info):
    """Real UTF-8 name of a zip entry (Keynote writes UTF-8 without the flag)."""
    if info.flag_bits & 0x800:
        return info.filename
    try:
        return info.filename.encode('cp437').decode('utf-8')
    except UnicodeError:
        return info.filename


def read_entries(path):
    """{名称: bytes}。支持两种 .key：单文件（zip）和包格式（文件夹，内含 Index.zip + Data/）。"""
    path = path.rstrip('/')
    if os.path.isdir(path):
        z = zipfile.ZipFile(os.path.join(path, 'Index.zip'))
        ents = {entry_name(i): z.read(i) for i in z.infolist()}
        for root, _, fs in os.walk(os.path.join(path, 'Data')):
            for f in fs:
                ents['Data/' + f] = open(os.path.join(root, f), 'rb').read()
        return ents
    z = zipfile.ZipFile(path)
    return {entry_name(i): z.read(i) for i in z.infolist()}


def _zip_cli(zpath, files):
    with tempfile.TemporaryDirectory() as td:
        for name, data in files.items():
            p = os.path.join(td, name)
            os.makedirs(os.path.dirname(p), exist_ok=True)
            open(p, 'wb').write(data)
        subprocess.run(['zip', '-q', '-0', '-X', os.path.abspath(zpath)] + list(files), cwd=td, check=True)


def zip_update(src, out, files):
    """复制 src→out，再用 zip 命令行逐条替换/新增条目 {名称: bytes}（不重打整包）。"""
    src, out = src.rstrip('/'), out.rstrip('/')
    if os.path.isdir(src):
        if os.path.abspath(src) != os.path.abspath(out):
            shutil.copytree(src, out)
        idx = {k: v for k, v in files.items() if not k.startswith('Data/')}
        if idx:
            _zip_cli(os.path.join(out, 'Index.zip'), idx)
        for k, v in files.items():
            if k.startswith('Data/'):
                open(os.path.join(out, k), 'wb').write(v)
        return
    if os.path.abspath(src) != os.path.abspath(out):
        shutil.copyfile(src, out)
    _zip_cli(out, files)


def load_iwa_bytes(raw):
    with tempfile.NamedTemporaryFile(suffix='.iwa', delete=False) as f:
        f.write(raw); p = f.name
    try:
        return read_iwa(p)
    finally:
        os.remove(p)


def dump_iwa_bytes(data):
    with tempfile.NamedTemporaryFile(suffix='.iwa', delete=False) as f:
        p = f.name
    try:
        write_iwa(p, data)
        assert read_iwa(p) == data
        return open(p, 'rb').read()
    finally:
        os.remove(p)


# ---------- chart grid ----------
def find_grids(b, min_rows=1):
    out = []
    for lp, ais, aie, msgs in objects(b):
        for mi, ps, pe in msgs:
            for f in fields(b, ps, pe):
                if f[0] == 10000 and f[1] == 2:
                    for g in fields(b, f[3], f[4]):
                        if g[0] == 7 and g[1] == 2:
                            try:
                                fl = list(fields(b, g[3], g[4]))
                            except ValueError:
                                continue
                            if sum(1 for x in fl if x[0] == 3) >= min_rows and any(x[0] == 2 for x in fl):
                                out.append(dict(lp=lp, ais=ais, aie=aie, mi=mi, ps=ps, pe=pe, f=f, g=g, fl=fl))
    return out


def grid_table(b, G):
    fl = G['fl']
    names = [b[x[3]:x[4]].decode('utf-8', 'replace') for x in fl if x[0] == 1]
    cols = [b[x[3]:x[4]].decode('utf-8', 'replace') for x in fl if x[0] == 2]
    rows = []
    for x in fl:
        if x[0] != 3:
            continue
        r = []
        for c in fields(b, x[3], x[4]):
            sub = list(fields(b, c[3], c[4]))
            r.append(struct.unpack('<d', b[sub[0][3]:sub[0][4]])[0] if sub and sub[0][1] == 1 else None)
        rows.append(r)
    return names, cols, rows


def _cell(v):
    inner = b'' if v is None else b'\x09' + struct.pack('<d', float(v))
    return b'\x0a' + ev(len(inner)) + inner


def _row(vals):
    body = b''.join(_cell(v) for v in vals)
    return b'\x1a' + ev(len(body)) + body


def _rewrap(b, ts, new_val):
    _, te = rv(b, ts)
    return b[ts:te] + ev(len(new_val)) + new_val


def rebuild_grid(b, G, names, rows):
    fl = G['fl']
    order = [x[0] for x in fl]
    if order != sorted(order):
        raise SystemExit('图表网格字段顺序异常，停止（请改用 Keynote 手动编辑）')
    old_n = sum(1 for x in fl if x[0] == 3)
    out = bytearray(); d1 = d3 = False
    for x in fl:
        if x[0] == 1:
            if not d1:
                for n in names:
                    nb = n.encode('utf-8'); out += b'\x0a' + ev(len(nb)) + nb
                d1 = True
        elif x[0] == 3:
            if not d3:
                for r in rows:
                    out += _row(r)
                d3 = True
        elif x[0] == 4:
            if len(names) == old_n:
                out += b[x[2]:x[4]]           # 行数不变：保留原行 ID
            else:                              # 行数变：重建行 ID，保留列 ID
                body = bytearray()
                for k in range(len(names)):
                    u = str(uuid.uuid4()).upper().encode()
                    ent = b'\x0a' + ev(len(u)) + u + b'\x10' + ev(k)
                    body += b'\x0a' + ev(len(ent)) + ent
                for y in fields(b, x[3], x[4]):
                    if y[0] != 1:
                        body += b[y[2]:y[4]]
                out += b'\x22' + ev(len(body)) + bytes(body)
        else:
            out += b[x[2]:x[4]]
    grid = bytes(out)
    f, g, ps, pe = G['f'], G['g'], G['ps'], G['pe']
    new10k = b[f[3]:g[2]] + _rewrap(b, g[2], grid) + b[g[4]:f[4]]
    payload = b[ps:f[2]] + _rewrap(b, f[2], new10k) + b[f[4]:pe]
    mts, mvs, mve = G['mi']
    if ps != G['aie']:
        raise SystemExit('该对象含多段消息，暂不支持，停止')
    mi_new = bytearray()
    for x in fields(b, mvs, mve):
        mi_new += (b[x[2]:x[3]] + ev(len(payload))) if x[0] == 3 else b[x[2]:x[4]]
    ai_new = b[G['ais']:mts] + _rewrap(b, mts, bytes(mi_new)) + b[mve:G['aie']]
    nb = b[:G['lp']] + ev(len(ai_new)) + ai_new + payload + b[pe:]
    for _ in objects(nb):  # 结构自检
        pass
    return nb


# ---------- commands ----------
def slide_entries(ents, only=None):
    for n in sorted(ents):
        if n.startswith('Index/Slide') and n.endswith('.iwa') and (only is None or n.endswith(only)):
            yield n


def cmd_find(a):
    ents = read_entries(a.key)
    for n in slide_entries(ents):
        if a.text.encode('utf-8') in load_iwa_bytes(ents[n]):
            print(n)


def cmd_charts(a):
    ents = read_entries(a.key)
    for n in slide_entries(ents, a.slide_file):
        b = load_iwa_bytes(ents[n])
        for k, G in enumerate(find_grids(b)):
            names, cols, rows = grid_table(b, G)
            print(f'{n} grid#{k}: 系列={cols} 行数={len(rows)} 首行={names[0] if names else ""}{rows[0] if rows else ""} 末行={names[-1] if names else ""}{rows[-1] if rows else ""}')


def _pick(a):
    ents = read_entries(a.key)
    hits = [n for n in ents if n.endswith(a.slide_file)]
    if len(hits) != 1:
        raise SystemExit(f'--slide-file 匹配到 {hits}')
    b = load_iwa_bytes(ents[hits[0]])
    grids = find_grids(b)
    if a.grid >= len(grids):
        raise SystemExit(f'该页只有 {len(grids)} 个图表')
    return ents, hits[0], b, grids[a.grid]


def cmd_get(a):
    _, n, b, G = _pick(a)
    names, cols, rows = grid_table(b, G)
    f = open(a.out, 'w', newline='', encoding='utf-8') if a.out else sys.stdout
    w = csv.writer(f)
    w.writerow([''] + cols)
    for nm, r in zip(names, rows):
        w.writerow([nm] + ['' if v is None else repr(v) for v in r])
    if a.out:
        f.close(); print(f'{n}: {len(rows)} 行 → {a.out}')


def cmd_set(a):
    if os.path.abspath(a.out) == os.path.abspath(a.key):
        raise SystemExit('--out 必须是新文件（先在副本上改）')
    ents, n, b, G = _pick(a)
    _, cols, _ = grid_table(b, G)
    rows = list(csv.reader(open(a.csv, encoding='utf-8-sig')))
    head, body = rows[0], [r for r in rows[1:] if r and r[0].strip()]
    if [h.strip() for h in head[1:]] != cols:
        raise SystemExit(f'CSV 系列名 {head[1:]} 与原图 {cols} 不一致')
    names = [r[0].strip() for r in body]
    vals = [[(float(v) if v.strip() else None) for v in (r[1:] + [''] * len(cols))[:len(cols)]] for r in body]
    nb = rebuild_grid(b, G, names, vals)
    zip_update(a.key, a.out, {n: dump_iwa_bytes(nb)})
    print(f'{n}: 写入 {len(names)} 行（{names[0]} → {names[-1]}）→ {a.out}')


def cmd_text(a):
    if os.path.abspath(a.out) == os.path.abspath(a.key):
        raise SystemExit('--out 必须是新文件')
    old, new = a.old.encode('utf-8'), a.new.encode('utf-8')
    if len(old) != len(new):
        raise SystemExit('只支持等长替换；长度不同请用 AppleScript 改文本框')
    ents = read_entries(a.key)
    hits = [k for k in ents if k.endswith(a.slide_file)]
    b = load_iwa_bytes(ents[hits[0]])
    c = b.count(old)
    if c != 1:
        raise SystemExit(f'「{a.old}」出现 {c} 次，需唯一')
    zip_update(a.key, a.out, {hits[0]: dump_iwa_bytes(b.replace(old, new))})
    print('ok')


def referenced_missing(ents):
    data = [k[5:] for k in ents if k.startswith('Data/')]
    m = load_iwa_bytes(ents['Index/Metadata.iwa'])
    refs = set(x.decode('utf-8', 'replace') for x in re.findall(rb'[^\x00-\x1f"]{1,160}?-\d{3,7}\.[A-Za-z0-9]{2,5}', m))
    return sorted(r for r in refs if not any(d in r for d in data))


def cmd_verify(a):
    A, B = read_entries(a.before), read_entries(a.after)
    print('删除:', sorted(k for k in A if k not in B))
    print('新增:', sorted(k for k in B if k not in A))
    print('修改:', sorted(k for k in A if k in B and A[k] != B[k]))
    miss = referenced_missing(B)
    print('被引用但缺失的图片/文件:', miss or '无')
    if miss:
        sys.exit(1)


def cmd_restore(a):
    cur, bak = read_entries(a.key), read_entries(a.backup)
    m = load_iwa_bytes(cur['Index/Metadata.iwa'])
    add = {k: v for k, v in bak.items() if k.startswith('Data/') and k not in cur and k[5:].encode('utf-8') in m}
    if not add:
        print('没有需要补回的文件'); return
    zip_update(a.key, a.out, add)
    print('补回:', sorted(add))


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    s = p.add_subparsers(dest='cmd', required=True)
    x = s.add_parser('find'); x.add_argument('key'); x.add_argument('text'); x.set_defaults(fn=cmd_find)
    x = s.add_parser('charts'); x.add_argument('key'); x.add_argument('--slide-file'); x.set_defaults(fn=cmd_charts)
    for nm, fn in (('get', cmd_get), ('set', cmd_set)):
        x = s.add_parser(nm); x.add_argument('key'); x.add_argument('--slide-file', required=True)
        x.add_argument('--grid', type=int, default=0); x.add_argument('--out', required=(nm == 'set'))
        if nm == 'set':
            x.add_argument('--csv', required=True)
        x.set_defaults(fn=fn)
    x = s.add_parser('text'); x.add_argument('key'); x.add_argument('--slide-file', required=True)
    x.add_argument('--old', required=True); x.add_argument('--new', required=True); x.add_argument('--out', required=True)
    x.set_defaults(fn=cmd_text)
    x = s.add_parser('verify'); x.add_argument('before'); x.add_argument('after'); x.set_defaults(fn=cmd_verify)
    x = s.add_parser('restore'); x.add_argument('key'); x.add_argument('--from', dest='backup', required=True)
    x.add_argument('--out', required=True); x.set_defaults(fn=cmd_restore)
    a = p.parse_args(); a.fn(a)


if __name__ == '__main__':
    main()
