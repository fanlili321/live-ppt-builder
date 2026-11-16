import sys,struct
sys.path.insert(0,'/Users/fanlili/.claude/skills/live-ppt-builder/scripts/keynote_iwa')
from keynote_tool import read_entries,load_iwa_bytes
from pb import fields,objects,rv,ev
key=sys.argv[1]; ents=read_entries(key)
def objs(name):
    b=load_iwa_bytes(ents[name]); out=[]
    for lp,ais,aie,msgs in objects(b):
        oid=None;types=[]
        for fn,wt,ts,vs,ve in fields(b,ais,aie):
            if fn==1: oid=rv(b,vs)[0]
            if fn==2:
                for f2,w2,t2,v2s,v2e in fields(b,vs,ve):
                    if f2==1: types.append(rv(b,v2s)[0])
        ps=msgs[0][1] if msgs else aie; pe=msgs[-1][2] if msgs else aie
        out.append(dict(id=oid,types=types,b=b,ps=ps,pe=pe))
    return out
def cols(o):
    b=o['b'];r=[]
    for i in range(o['ps'],o['pe']-15):
        if b[i]==0x1d and b[i+5]==0x25 and b[i+10]==0x2d:
            v=[struct.unpack('<f',b[i+1+5*k:i+5+5*k])[0] for k in range(3)]
            if all(0<=x<=1 for x in v): r.append('#%02X%02X%02X'%tuple(round(x*255) for x in v))
    return r
ss={o['id']:o for o in objs('Index/DocumentStylesheet.iwa')}
targets={i:(o['types'],cols(o)) for i,o in ss.items() if set(cols(o))&{'#E97C2D','#85A6DB','#6975A2','#929292'} and o['types'] and o['types'][0]!=0}
for sl in sys.argv[2:]:
    for o in objs(sl):
        b=o['b'];seg=b[o['ps']:o['pe']]
        hits=[]
        for i,(t,c) in targets.items():
            pat=b'\x08'+ev(i)
            if pat in seg: hits.append((i,t[0],c[:2]))
        if hits: print(sl,'obj',o['id'],'type',o['types'][:3],'->',sorted(hits))
if '--show' in sys.argv[2:3] or True:
    pass
