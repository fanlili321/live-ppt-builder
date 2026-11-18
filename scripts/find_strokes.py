import sys,struct
sys.path.insert(0,'/Users/fanlili/.claude/skills/live-ppt-builder/scripts/keynote_iwa')
import keynote_tool as k
from pb import rv,fields,ev,objects
def is_color(b,s,e):
    try: fl=list(fields(b,s,e))
    except Exception: return None
    d={f[0]:f for f in fl}
    if 1 in d and d[1][1]==0 and 3 in d and d[3][1]==5 and 4 in d and 5 in d:
        return tuple(round(struct.unpack('<f',b[d[i][3]:d[i][4]])[0],3) for i in (3,4,5))
    return None
def walk(b,s,e,path,out,depth=0):
    try: fl=list(fields(b,s,e))
    except Exception: return
    d={}
    for f in fl: d.setdefault(f[0],f)
    # stroke: field1 color msg, field2 f32 width
    if 1 in d and d[1][1]==2 and 2 in d and d[2][1]==5:
        c=is_color(b,d[1][3],d[1][4])
        if c is not None:
            out.append((path,struct.unpack('<f',b[d[2][3]:d[2][4]])[0],c,d[2][3]))
    if depth>12: return
    for f in fl:
        if f[1]==2 and f[4]-f[3]>2: walk(b,f[3],f[4],path+[f[0]],out,depth+1)
def scan(b):
    res=[]
    for lp,ais,aie,msgs in objects(b):
        ident=[rv(b,f[3])[0] for f in fields(b,ais,aie) if f[0]==1][0]
        types=[rv(b,g[3])[0] for f in fields(b,ais,aie) if f[0]==2 for g in fields(b,f[3],f[4]) if g[0]==1]
        for i,(_,ps,pe) in enumerate(msgs):
            out=[];walk(b,ps,pe,[],out)
            for o in out: res.append((ident,types[i] if i<len(types) else None)+o)
    return res
if __name__=='__main__':
    e=k.read_entries(sys.argv[1])
    for n in sys.argv[2:]:
        b=k.load_iwa_bytes(e[n])
        for r in scan(b): print(n,r)
