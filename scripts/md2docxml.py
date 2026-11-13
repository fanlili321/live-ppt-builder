"""把过程记录 md 转成飞书 DocxXML 片段（一节）。用法：md2docxml.py in.md 节标题 > out.xml"""
import sys, re, html
src, h1 = sys.argv[1], sys.argv[2]
def inl(t):
    t = html.escape(t, quote=False)
    t = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', t)
    t = re.sub(r'`([^`]+)`', r'\1', t)
    t = re.sub(r'\[([^\]]+)\]\([^)]+\)', r'\1', t)
    return t
lines = open(src, encoding='utf-8').read().splitlines()
out = [f'<h1>{inl(h1)}</h1>']
i = 0
# 跳过 md 一级标题
while i < len(lines) and not lines[i].startswith('## ') and not lines[i].startswith('- '): i += 1
def flush_list(items):
    # items: [(indent, text)]
    res, stack = [], []
    for ind, txt in items:
        lvl = ind // 2
        while len(stack) > lvl + 1: res.append('</li></ul>'); stack.pop()
        if len(stack) == lvl + 1: res.append('</li><li>' + f'<p>{inl(txt)}</p>')
        else:
            res.append('<ul><li>' + f'<p>{inl(txt)}</p>'); stack.append(lvl)
    while stack: res.append('</li></ul>'); stack.pop()
    return ''.join(res)
while i < len(lines):
    l = lines[i]
    if l.startswith('## '):
        out.append(f'<h2>{inl(l[3:])}</h2>'); i += 1
    elif l.startswith('|'):
        rows = []
        while i < len(lines) and lines[i].startswith('|'):
            cells = [c.strip() for c in lines[i].strip().strip('|').split('|')]
            if not all(re.fullmatch(r':?-+:?', c) for c in cells): rows.append(cells)
            i += 1
        n = len(rows[0])
        t = '<table><colgroup>' + '<col/>' * n + '</colgroup><tbody>'
        for k, r in enumerate(rows):
            t += '<tr>' + ''.join(f'<td><p>{"<b>" if k == 0 else ""}{inl(c)}{"</b>" if k == 0 else ""}</p></td>' for c in r) + '</tr>'
        out.append(t + '</tbody></table>')
    elif re.match(r'\s*(- |\d+\. )', l):
        items = []
        while i < len(lines) and re.match(r'\s*(- |\d+\. )', lines[i]):
            m = re.match(r'(\s*)(- |\d+\. )(.*)', lines[i]); items.append((len(m.group(1)), m.group(3))); i += 1
        out.append(flush_list(items))
    elif l.strip():
        out.append(f'<p>{inl(l)}</p>'); i += 1
    else:
        i += 1
print('\n'.join(out))
