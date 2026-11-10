# -*- coding: utf-8 -*-
import sys,json; sys.path.insert(0,'/Users/fanlili/Downloads/live-ppt-tool')
import live_helpers as hl
S,E='2020-01-01','2026-06-12'
# name:(price_code, totalreturn_code or None)  港股部分代码未确认→None占位
M={
 '中证全指':('000985.CSI','H00985.CSI'),'沪深300':('000300.SH','H00300.CSI'),
 '恒生指数':('HSI.HI',None),'H股指数':('HSCEI.HI',None),
 '中证500':('000905.SH','H00905.CSI'),'中证1000':('000852.SH','H00852.CSI'),
 '创业板指':('399006.SZ',None),'科创50':('000688.SH',None),
 '中证红利':('000922.CSI','H00922.CSI'),'中证红利低波':('H30269.CSI',None),
 '恒生红利低波动':(None,None),'中证A500':('000510.SH','H00510.CSI'),
 '消费50':('931139.CSI',None),'恒生消费':(None,None),
 '医药100':('000978.SH',None),'恒生医疗保健':(None,None),
 '港股科技':('HSTECH.HI',None),
}
out={}
for n,(pc,tc) in M.items():
    rec={}
    for tag,code in [('price',pc),('tr',tc)]:
        if not code: rec[tag]=None; continue
        try:
            d=hl.kl(code,S,E,'W'); rec[tag]={'code':code,'n':len(d),'data':d} if d else None
            print(f"{'OK ' if d else 'EMP'} {n}/{tag} {code} {len(d)}",flush=True)
        except Exception as ex:
            rec[tag]=None; print(f"ERR {n}/{tag} {code} {str(ex)[:30]}",flush=True)
    out[n]=rec
json.dump(out,open('cache_446.json','w'),ensure_ascii=False)
print("DONE",sum(1 for n in out if out[n].get('price')),"指数有价格线")
