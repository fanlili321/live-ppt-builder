"""写图表数据并删掉一个系列列：python3 set_drop_col.py key slide.iwa grid csv drop_idx out
保留其余列的列ID（及其 index 值顺延），列名按 CSV 表头。"""
import os
import sys,csv,uuid
sys.path.insert(0,os.path.join(os.path.dirname(os.path.abspath(__file__)),'keynote_iwa'))
import keynote_tool as k
from pb import ev
key,sf,gi,cf,drop,out=sys.argv[1],sys.argv[2],int(sys.argv[3]),sys.argv[4],int(sys.argv[5]),sys.argv[6]
ents=k.read_entries(key);n=[x for x in ents if x.endswith(sf)][0]
b=k.load_iwa_bytes(ents[n]);G=k.find_grids(b)[gi]
rows=list(csv.reader(open(cf)));head=[h.strip() for h in rows[0][1:]];body=[r for r in rows[1:] if r and r[0]]
names=[r[0] for r in body];vals=[[float(v) if v else None for v in r[1:]] for r in body]
fl=G['fl'];out_b=bytearray();d1=d2=d3=False
for x in fl:
    if x[0]==2:
        if not d2:
            for h in head:
                hb=h.encode();out_b+=b'\x12'+ev(len(hb))+hb
            d2=True
    elif x[0]==1:
        if not d1:
            for nm in names:
                nb=nm.encode();out_b+=b'\x0a'+ev(len(nb))+nb
            d1=True
    elif x[0]==3:
        if not d3:
            for r in vals: out_b+=k._row(r)
            d3=True
    elif x[0]==4:
        body2=bytearray()
        for i in range(len(names)):
            u=str(uuid.uuid4()).upper().encode();ent=b'\x0a'+ev(len(u))+u+b'\x10'+ev(i)
            body2+=b'\x0a'+ev(len(ent))+ent
        ci=0
        for y in k.fields(b,x[3],x[4]):
            if y[0]==1: continue
            assert y[0]==2,y
            sub=list(k.fields(b,y[3],y[4]));uid=b[sub[0][3]:sub[0][4]]
            if ci==drop: ci+=1; continue
            newi=ci if ci<drop else ci-1
            ent=b'\x0a'+ev(len(uid))+uid+b'\x10'+ev(newi)
            body2+=b'\x12'+ev(len(ent))+ent
            ci+=1
        out_b+=b'\x22'+ev(len(body2))+bytes(body2)
    else:
        out_b+=b[x[2]:x[4]]
grid=bytes(out_b)
f,g,ps,pe=G['f'],G['g'],G['ps'],G['pe']
new10k=b[f[3]:g[2]]+k._rewrap(b,g[2],grid)+b[g[4]:f[4]]
payload=b[ps:f[2]]+k._rewrap(b,f[2],new10k)+b[f[4]:pe]
mts,mvs,mve=G['mi']
mi_new=bytearray()
for x in k.fields(b,mvs,mve):
    mi_new+=(b[x[2]:x[3]]+ev(len(payload))) if x[0]==3 else b[x[2]:x[4]]
ai_new=b[G['ais']:mts]+k._rewrap(b,mts,bytes(mi_new))+b[mve:G['aie']]
nb=b[:G['lp']]+ev(len(ai_new))+ai_new+payload+b[pe:]
for _ in k.objects(nb): pass
k.zip_update(key,out,{n:k.dump_iwa_bytes(nb)})
print('ok',len(names),head)
