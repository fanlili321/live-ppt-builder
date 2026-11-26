"""490 期终审：工具箱重新取数，与在线文档图表逐点对照，并重算表格数字（Wind 不可用时也用它）。
用法（在 output/NNN_claude输出 下）：
  1. cp -R 在线文档 data/终审_online.key；对每张图 keynote_tool.py get ... --grid N --out data/audit/pN.csv
  2. 改下面 comp(...) / 表格期望值 / P6P7 估值文件名为本期配置
  3. python3 scripts/final_audit_例490.py → 每项「一致/不一致」，结果存 data/audit/result.json
"""
import csv,json,os,subprocess,datetime as dt
D=os.path.abspath('../../.agents/skills/live-ppt-builder/scripts/dbms_api.py'); A='data/audit'
def fetch(code,start,end):
    fn=f'{A}/src_{code}_{start}_{end}.json'
    if os.path.exists(fn): return json.load(open(fn))
    out={}; s=dt.date.fromisoformat(start); e=dt.date.fromisoformat(end)
    while s<=e:
        t=min(dt.date(s.year+2,12,31),e)
        q=json.dumps({"codes":[code],"beginDate":str(s),"endDate":str(t),"frequency":"daily","fields":["close"]})
        d=None
        for _ in range(3):
            r=subprocess.run(['python3',D,'index-query',q],capture_output=True,text=True,timeout=170)
            try: d=json.loads(r.stdout)['data']; break
            except Exception: d=None
        if d is None: raise SystemExit(f'取数失败 {code} {s}')
        for day,row in zip(d['dates'],d['rows']):
            v=list(row.values())[0]
            if v is not None: out[day]=v
        s=t+dt.timedelta(days=1)
    json.dump(out,open(fn,'w')); return out
def deck(n):
    r=list(csv.reader(open(f'{A}/{n}.csv'))); h=r[0][1:]
    return h,{x[0]:[float(v) if v not in ('','None') else None for v in x[1:]] for x in r[1:]}
def mdd(s):
    ks=sorted(s); pk=ks[0]; best=(0,None,None)
    for k in ks:
        if s[k]>s[pk]: pk=k
        dd=s[k]/s[pk]-1
        if dd<best[0]: best=(dd,pk,k)
    return best
cal=None
res=[]
def chk(page,name,ok,detail): res.append((page,name,'一致' if ok else '不一致',detail)); print(page,name,'OK' if ok else 'XX',detail)
END='2026-11-25'
S={}
for c in ['H00985.CSI','399606.SZ','CN2673.CNI','H20269.CSI','H00922.CSI','000985.CSI','399006.SZ']:
    S[c]=fetch(c,'2010-05-31',END)
cal=sorted(S['399006.SZ'])
S['CN2673.CNI']={k:v for k,v in S['CN2673.CNI'].items() if k in set(cal)}
for c in ['399606.SZ','CN2673.CNI']: S[c].setdefault('2010-05-31',1000.0)
def comp(page,csvn,codes,base):
    h,dk=deck(csvn); worst=0; n=0
    for name,code in zip(h,codes):
        s=S[code]; b=s[base]
        for day,vals in dk.items():
            v=vals[h.index(name)]
            if v is None or day not in s: continue
            worst=max(worst,abs(v-(s[day]/b-1))); n+=1
    chk(page,'+'.join(h),worst<0.0006,f'{n} 个点，最大偏差 {worst:.5f}（图上保留 4 位小数）')
comp('P8','p8',['H00985.CSI','399606.SZ','CN2673.CNI'],'2010-05-31')
comp('P9','p9',['399606.SZ','CN2673.CNI'],'2025-12-31')
comp('P12','p12',['H20269.CSI','399606.SZ','CN2673.CNI'],'2024-08-30')
comp('P13','p13',['H20269.CSI','399606.SZ','CN2673.CNI'],'2026-06-30')
comp('P19','p19',['H00922.CSI','399606.SZ'],'2018-12-28')
# P8 table
for code,ann,md in [('H00985.CSI',4.08,-55.78),('399606.SZ',5.56,-69.23),('CN2673.CNI',5.46,-74.59)]:
    s={k:v for k,v in S[code].items() if k<=END}; y=(dt.date.fromisoformat(END)-dt.date(2010,5,31)).days/365.25
    a=((s[END]/s['2010-05-31'])**(1/y)-1)*100; m=mdd(s)
    chk('P8表',code,abs(a-ann)<0.006 and abs(m[0]*100-md)<0.006,f'年化 {a:.2f}%（页面 {ann}%），最大回撤 {m[0]*100:.2f}% {m[1]}→{m[2]}（页面 {md}%）')
for code,cum,md in [('399606.SZ',-31.25,-49.93),('CN2673.CNI',-35.46,-53.76)]:
    s={k:v for k,v in S[code].items() if '2025-12-31'<=k<=END}; c=(s[END]/s['2025-12-31']-1)*100; m=mdd(s)
    chk('P9表',code,abs(c-cum)<0.006 and abs(m[0]*100-md)<0.006,f'累计 {c:.2f}%（页面 {cum}%），回撤 {m[0]*100:.2f}% {m[1]}→{m[2]}（页面 {md}%）')
ann19={'H00922.CSI':[20.88,8.18,18.19,-0.37,6.34,18.76,3.76,2.69],'399606.SZ':[44.81,65.99,12.41,-29.03,-18.77,14.83,51.56,-31.25]}
for code,vals in ann19.items():
    s=S[code]; got=[]
    for y in range(2019,2027):
        prev=max(k for k in s if k<f'{y}-01-01'); last=max(k for k in s if k<=min(f'{y}-12-31',END)); got.append(round((s[last]/s[prev]-1)*100,2))
    chk('P19表',code,all(abs(a-b)<0.006 for a,b in zip(got,vals)),f'重算 {got}')
# P14
mg=json.load(open('data/margin.json')) if os.path.exists('data/margin.json') else None
h,dk=deck('p14c'); s=S['000985.CSI']; w=max(abs(v[0]-s[d]) for d,v in dk.items() if d in s and v[0] is not None)
chk('P14','中证全指',w<0.01,f'{len(dk)} 个点，最大偏差 {w:.3f}')
# P6/P7
for p,n in [('P6','创业板'),('P7','创业板50')]:
    v=json.load(open(f'data/v3/{n}_all.json')); h,dk=deck(p.lower())
    bad=[d for d,x in dk.items() if abs(x[0]-v[d][0])>0.001 or abs(x[1]-v[d][1])>0.001]
    chk(p,n+' PE/PB',not bad and len(dk)==4004,f'{len(dk)} 行，与工具箱不一致 {len(bad)}')
json.dump(res,open(f'{A}/result.json','w'),ensure_ascii=False,indent=1)
