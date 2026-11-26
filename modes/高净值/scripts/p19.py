import json,pandas as pd
from PIL import Image,ImageDraw,ImageFont,ImageFilter
F='/System/Library/AssetsV2/com_apple_MobileAsset_Font8/86ba2c91f017a3749571a82f2c6d890ac7ffb2fb.asset/AssetData/PingFang.ttc'
df=pd.read_excel('../src/nav/lark-im-resources/file_v3_0017o_23a032a2-d60d-4dfa-ae1a-638e8260a2fg.xls')
df['d']=pd.to_datetime(df['净值日期'].astype(str)).dt.strftime('%Y-%m-%d')
nav=dict(zip(df.d,df['复权净值'].astype(float)))
hs=json.load(open('hs300.json'))
dates=sorted(d for d in nav if d>='2021-06-21' and d<='2026-11-23' and d in hs)
b=hs['2021-06-21']
rows=[(d,nav[d],hs[d]/b) for d in dates]
pd.DataFrame(rows,columns=['date','strategy','hs300']).assign(excess=lambda x:x.strategy-x.hs300).to_csv('p19_series.csv',index=False)
print(len(rows),rows[0],rows[-1],'excess end',rows[-1][1]-rows[-1][2])
K=2; W,H=1364*K,690*K
bg=(250,248,243); im=Image.new('RGB',(W,H),bg); dr=ImageDraw.Draw(im)
x0,x1=105*K,1260*K; y0=378*K
Y=lambda v: y0-(v-1)*452*K
YE=lambda e: y0-e*565*K
n=len(rows)-1; X=lambda i: x0+(x1-x0)*i/n
f_axis=ImageFont.truetype(F,22*K,index=0)
# grid + axis labels
for v,lab,r in [(1.5,'1.50','40%'),(1.0,'1.00','0%'),(0.5,'0.50','-40%')]:
    y=Y(v); dr.line([(95*K,y),(1270*K,y)],fill=(232,230,226),width=K)
    dr.text((35*K,y),lab,font=f_axis,fill=(150,150,150),anchor='lm')
    dr.text((1278*K,y),r,font=f_axis,fill=(150,150,150),anchor='lm')
# excess area
poly=[(X(i),YE(s-h)) for i,(d,s,h) in enumerate(rows)]
dr.polygon([(X(0),y0)]+poly+[(X(n),y0)],fill=(243,223,159))
dr.line([(X(i),Y(h)) for i,(d,s,h) in enumerate(rows)],fill=(187,195,202),width=5*K,joint='curve')
dr.line([(X(i),Y(s)) for i,(d,s,h) in enumerate(rows)],fill=(220,153,88),width=7*K,joint='curve')
# blur plot region
box=(0,110*K,W,632*K)
reg=im.crop(box).filter(ImageFilter.GaussianBlur(6*K)); im.paste(reg,box[:2])
dr=ImageDraw.Draw(im)
# legend
fl=ImageFont.truetype(F,30*K,index=0)
dr.line([(124*K,54*K),(200*K,54*K)],fill=(220,153,88),width=7*K); dr.text((222*K,54*K),'某主动优选私募策略',font=fl,fill=(23,23,23),anchor='lm')
dr.line([(555*K,54*K),(630*K,54*K)],fill=(187,195,202),width=5*K); dr.text((652*K,54*K),'沪深300',font=fl,fill=(23,23,23),anchor='lm')
dr.rectangle([(831*K,43*K),(900*K,66*K)],fill=(243,223,159)); dr.text((924*K,54*K),'超额（相对沪深300）',font=fl,fill=(23,23,23),anchor='lm')
fd=ImageFont.truetype(F,28*K,index=0)
dr.text((104*K,658*K),'2021-06-21',font=fd,fill=(83,84,84),anchor='lm')
dr.text((1260*K,658*K),'2026-11-23',font=fd,fill=(83,84,84),anchor='rm')
# border
dr.rounded_rectangle([(K,K),(W-1-K,H-1-K)],radius=14*K,outline=(238,200,133),width=3*K)
# outside corners transparent-ish -> paint slide bg white-ish
im.save('p19_1123.png')
