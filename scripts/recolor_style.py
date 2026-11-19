"""把指定形状样式的描边改色（DocumentStylesheet 里原地替换 float，不改长度）。
用法：recolor_style.py in.key out.key 样式id r g b [--dry-run]
  r g b 为 0-1 小数，如红色 0.9 0 0.055。--dry-run 只打印原色、粗细、虚线类型和使用该样式的所有形状。
注意：样式可能被多页共用，先 --dry-run 看「使用者」，只在本页使用时再改（488 期 P14 虚线框蓝→红）。
先在 Keynote 里关闭文件；改完 keynote_tool.py verify。"""
import sys, os, struct
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, 'keynote_iwa'))
sys.path.insert(0, HERE)
import keynote_tool as kt
from pb import fields, rv
from set_frame_strokes import objs, sub
a=[x for x in sys.argv[1:] if not x.startswith('--')]; dry='--dry-run' in sys.argv
src,out,sid=a[0],a[1],int(a[2]); rgb=tuple(map(float,a[3:6]))
E=kt.read_entries(src)
st=bytearray(kt.load_iwa_bytes(E['Index/DocumentStylesheet.iwa']))
STY={oid:(ps,pe) for oid,t,ps,pe in objs(bytes(st))}
ps,pe=STY[sid]; f1=sub(st,ps,pe,1); p=sub(st,f1[3],f1[4],11); sk=sub(st,p[3],p[4],2)
col=sub(st,sk[3],sk[4],1); wf=sub(st,sk[3],sk[4],2); pat=sub(st,sk[3],sk[4],6)
fs=[sub(st,col[3],col[4],i) for i in (3,4,5)]
print('原色',[round(struct.unpack('<f',st[f[3]:f[4]])[0],3) for f in fs],'粗细',struct.unpack('<f',st[wf[3]:wf[4]])[0],'虚线类型',rv(st,sub(st,pat[3],pat[4],1)[3])[0] if pat else None)
# 使用该样式的形状
for k in sorted(E):
    if not (k.startswith('Index/Slide') and k.endswith('.iwa')): continue
    b=kt.load_iwa_bytes(E[k])
    for oid,t,ps2,pe2 in objs(b):
        if t!=2011: continue
        g1=sub(b,ps2,pe2,1); g12=sub(b,g1[3],g1[4],2)
        if rv(b,sub(b,g12[3],g12[4],1)[3])[0]!=sid: continue
        g=sub(b,sub(b,g1[3],g1[4],1)[3],sub(b,g1[3],g1[4],1)[4],1); size=sub(b,g[3],g[4],2)
        print('使用者',k,oid,[round(struct.unpack('<f',b[x[3]:x[4]])[0]) for x in fields(b,size[3],size[4])])
if dry: sys.exit()
for f,v in zip(fs,rgb): st[f[3]:f[4]]=struct.pack('<f',v)
kt.zip_update(src,out,{'Index/DocumentStylesheet.iwa':kt.dump_iwa_bytes(bytes(st))})
print('→',out)
