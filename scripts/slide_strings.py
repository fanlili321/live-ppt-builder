"""比对两版 .key 每页新增的中文字符串（找用户新加的批注/便签、改动的文字）。
用法：slide_strings.py 旧.key 新.key
批注也可能是形状里的文字：再用 AppleScript 列出每页带文字的 shape 交叉确认。"""
import sys,zipfile,re,io
import os
sys.path.insert(0,os.path.join(os.path.dirname(os.path.abspath(__file__)),'keynote_iwa'))
import iwa
def strings(key):
    z=zipfile.ZipFile(key); out={}
    for n in z.namelist():
        if not n.startswith('Index/') or not n.endswith('.iwa'): continue
        raw=z.read(n)
        import tempfile,os
        fd,p=tempfile.mkstemp(suffix='.iwa'); os.close(fd); open(p,'wb').write(raw)
        try: d=iwa.read_iwa(p)
        finally: os.remove(p)
        s=set(m.decode('utf-8','ignore') for m in re.findall(rb'(?:[\xe4-\xe9][\x80-\xbf]{2}|[\x20-\x7e]){4,}',d) if re.search(rb'[\xe4-\xe9]',m))
        out[n]=s
    return out
a=strings(sys.argv[1]); b=strings(sys.argv[2])
for n in sorted(b):
    new=b[n]-a.get(n,set())
    if new: print('==',n); [print('  +',x[:200]) for x in sorted(new)]
