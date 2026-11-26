"""把指定幻灯片里图表系列样式的折线描边宽度改成 W：python3 set_line_width.py in.key out.key W field Slide-xxx.iwa ...
系列样式在 DocumentStylesheet.iwa（type 5028），描边在 10000 下的 field（默认 86）。"""
import os
import sys,struct
sys.path.insert(0,os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0,os.path.join(os.path.dirname(os.path.abspath(__file__)),'keynote_iwa'))
import keynote_tool as k, find_strokes as fs
from pb import ev
src,dst,W,field=sys.argv[1],sys.argv[2],float(sys.argv[3]),int(sys.argv[4]);slides=sys.argv[5:]
e=k.read_entries(src);ds=bytearray(k.load_iwa_bytes(e['Index/DocumentStylesheet.iwa']))
res=fs.scan(bytes(ds))
sl=[k.load_iwa_bytes(e['Index/'+s]) for s in slides]
n=0;seen=set()
for ident,t,path,w,c,off in res:
    if t!=5028 or path!=[10000,field]: continue
    if not any(b.count(ev(ident)) for b in sl): continue
    if off in seen: continue
    seen.add(off)
    ds[off:off+4]=struct.pack('<f',W);n+=1;print('set',ident,w,'->',W,c)
k.zip_update(src,dst,{'Index/DocumentStylesheet.iwa':k.dump_iwa_bytes(bytes(ds))});print('changed',n)
