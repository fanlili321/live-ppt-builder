# 对调某页图表里两个系列样式的引用（等长 varint 才换），只改这一个 slide iwa
import os
import sys
sys.path.insert(0,os.path.join(os.path.dirname(os.path.abspath(__file__)),'keynote_iwa'))
from keynote_tool import read_entries,load_iwa_bytes,dump_iwa_bytes,zip_update
from pb import ev
key,slide,a,b_,out=sys.argv[1],sys.argv[2],int(sys.argv[3]),int(sys.argv[4]),sys.argv[5]
ents=read_entries(key); b=bytearray(load_iwa_bytes(ents[slide]))
pa,pb_=b'\x08'+ev(a),b'\x08'+ev(b_); assert len(pa)==len(pb_)
ia=[i for i in range(len(b)) if b[i:i+len(pa)]==pa]; ib=[i for i in range(len(b)) if b[i:i+len(pb_)]==pb_]
print('refs',ia,ib); assert len(ia)==1 and len(ib)==1
b[ia[0]:ia[0]+len(pa)]=pb_; b[ib[0]:ib[0]+len(pb_)]=pa
zip_update(key,out,{slide:dump_iwa_bytes(bytes(b))}); print('ok',out)
