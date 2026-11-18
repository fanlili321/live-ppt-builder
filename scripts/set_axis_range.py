"""设定图表值轴 min/max：python3 set_axis_range.py in.key out.key Slide-x.iwa grid axisIdx min max  (min/max 用 'none' 跳过)
值轴 nonstyle（type 5027）引用在图表信息的 field 14；min=field 18.1，max=field 17.1（double）。"""
import sys,struct
sys.path.insert(0,'/Users/fanlili/.claude/skills/live-ppt-builder/scripts/keynote_iwa')
import keynote_tool as k
from pb import rv,fields,ev
src,dst,sf,gi,ai,mn,mx=sys.argv[1:8];gi=int(gi);ai=int(ai)
e=k.read_entries(src);sn=[n for n in e if n.endswith(sf)][0]
b=k.load_iwa_bytes(e[sn]);G=k.find_grids(b)[gi];f=G['f']
refs=[rv(b,list(fields(b,x[3],x[4]))[0][3])[0] for x in fields(b,f[3],f[4]) if x[0]==14]
target=refs[ai];print('axis obj',target,refs)
oc=k.load_iwa_bytes(e['Index/ObjectContainer.iwa'])
def dbl(fn,v):
    inner=b'\x09'+struct.pack('<d',v);return ev(fn<<3|2)+ev(len(inner))+inner
for lp,ais,aie,msgs in k.objects(oc):
    ident=[rv(oc,x[3])[0] for x in fields(oc,ais,aie) if x[0]==1][0]
    if ident!=target: continue
    (mts,mvs,mve),ps,pe=msgs[0]
    pay=oc[ps:pe]
    out=bytearray()
    for x in fields(oc,ps,pe):
        if x[0]==10000:
            body=bytearray()
            for y in fields(oc,x[3],x[4]):
                if (y[0]==18 and mn!='none') or (y[0]==17 and mx!='none'): continue
                body+=oc[y[2]:y[4]]
            if mx!='none': body+=dbl(17,float(mx))
            if mn!='none': body+=dbl(18,float(mn))
            out+=oc[x[2]:x[3]-len(ev(x[4]-x[3]))] if False else b''
            out+=ev(10000<<3|2)+ev(len(body))+body
        else:
            out+=oc[x[2]:x[4]]
    newpay=bytes(out)
    mi=bytearray()
    for x in fields(oc,mvs,mve):
        mi+=(oc[x[2]:x[3]]+ev(len(newpay))) if x[0]==3 else oc[x[2]:x[4]]
    ai_new=oc[ais:mts]+k._rewrap(oc,mts,bytes(mi))+oc[mve:aie]
    noc=oc[:lp]+ev(len(ai_new))+ai_new+newpay+oc[pe:]
    for _ in k.objects(noc): pass
    k.zip_update(src,dst,{'Index/ObjectContainer.iwa':k.dump_iwa_bytes(noc)});print('ok');break
else: raise SystemExit('not found')
