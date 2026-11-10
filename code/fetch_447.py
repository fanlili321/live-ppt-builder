# -*- coding: utf-8 -*-
import sys,json; sys.path.insert(0,'/Users/fanlili/Downloads/live-ppt-tool')
import live_helpers as hl
out={}
# 对比图 2026初至今 日频
cmp={}
for n,c in [('中证红利全收益','H00922.CSI'),('中证全指自由现金流全收益','H11141.CSI'),('货币基金总指数','H11025.CSI')]:
    try:
        d=hl.kl(c,'2026-01-01','2026-06-12','D'); cmp[n]={'code':c,'data':d}
        print(f"OK cmp {n} {len(d)}",flush=True)
    except Exception as e: print("ERR",n,str(e)[:25],flush=True)
out['compare']=cmp
# 组合业绩 成立以来 (日频净值)
for key,name in [('active','螺丝钉金钉宝主动优选C'),('enh','螺丝钉金钉宝指数增强C')]:
    d=hl.combo(name,2); out[key]={'name':name,'data':d}
    print(f"OK combo {name} {len(d)}",flush=True)
json.dump(out,open('cache_447.json','w'),ensure_ascii=False)
print("DONE",flush=True)
