"""把指定段落样式(type 2022)的 alignment 设为 3（两端对齐），需要时插入字段并重算长度。
用法: python3 set_align.py in.key out.key 样式id[,样式id...] [对齐值=3]"""
import sys, os; sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'keynote_iwa'))
import keynote_tool as kt
from pb import fields, rv, ev, objects
src,out,ids=sys.argv[1],sys.argv[2],[int(x) for x in sys.argv[3].split(',')]
AL=int(sys.argv[4]) if len(sys.argv)>4 else 3
E=kt.read_entries(src); b=kt.load_iwa_bytes(E['Index/DocumentStylesheet.iwa'])
def new_payload(b,ps,pe):
    fl=list(fields(b,ps,pe)); o=bytearray(); done=False
    for fn,wt,ts,vs,ve in fl:
        if fn==12 and wt==2:
            sub=list(fields(b,vs,ve)); body=bytearray()
            for g in sub:
                if g[0]==1: body+=b'\x08'+ev(AL); done=True
                else:
                    if not done and g[0]>1: body+=b'\x08'+ev(AL); done=True
                    body+=b[g[2]:g[4]]
            if not done: body+=b'\x08'+ev(AL); done=True
            o+=b'\x62'+ev(len(body))+body
        else:
            if not done and fn>12: o+=b'\x62\x02\x08'+ev(AL); done=True
            o+=b[ts:ve]
    if not done: o+=b'\x62\x02\x08'+ev(AL)
    return bytes(o)
for target in ids:
    for lp,ais,aie,msgs in objects(b):
        oid=None
        for fn,wt,ts,vs,ve in fields(b,ais,aie):
            if fn==1: oid,_=rv(b,vs)
        if oid!=target: continue
        assert len(msgs)==1,'multi msg'
        (mts,mvs,mve),ps,pe=msgs[0]
        t=[rv(b,x[3])[0] for x in fields(b,mvs,mve) if x[0]==1][0]; assert t==2022,t
        pay=new_payload(b,ps,pe)
        mi=bytearray()
        for x in fields(b,mvs,mve):
            mi+=(b[x[2]:x[3]]+ev(len(pay))) if x[0]==3 else b[x[2]:x[4]]
        _,te=rv(b,mts)
        ai=b[ais:mts]+b[mts:te]+ev(len(mi))+bytes(mi)+b[mve:aie]
        b=b[:lp]+ev(len(ai))+ai+pay+b[pe:]
        print('set',target); break
    else: sys.exit('not found %d'%target)
for _ in objects(b): pass
kt.zip_update(src,out,{'Index/DocumentStylesheet.iwa':kt.dump_iwa_bytes(b)})
