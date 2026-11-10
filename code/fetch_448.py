# -*- coding: utf-8 -*-
import sys,json; sys.path.insert(0,'/Users/fanlili/Downloads/live-ppt-tool')
import live_helpers as hl
out={}
# CPPI案例：中证红利低波全收益 2021至今 按周
try:
    d=hl.kl('H30269.CSI','2021-01-01','2026-06-12','W'); out['hongli_lowvol']={'code':'H30269.CSI','data':d}
    print(f"OK 红利低波全收益 {len(d)}",flush=True)
except Exception as e: print("ERR 红利低波",str(e)[:25],flush=True)
# 90天 / 365天 组合 成立以来
for key,name in [('d90','螺丝钉银钉宝90天C'),('d365','螺丝钉银钉宝365天C')]:
    d=hl.combo(name,2); out[key]={'name':name,'data':d}
    print(f"OK {name} {len(d)}",flush=True)
json.dump(out,open('cache_448.json','w'),ensure_ascii=False)
print("DONE",flush=True)
