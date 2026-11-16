import sys,struct,collections
sys.path.insert(0,'/Users/fanlili/.claude/skills/live-ppt-builder/scripts/keynote_iwa')
from keynote_tool import read_entries,load_iwa_bytes,entry_name
key=sys.argv[1]; only=sys.argv[2:] 
ents=read_entries(key)
for name,raw in ents.items():

    if only and not any(o in name for o in only): continue
    if not name.endswith('.iwa'): continue
    b=load_iwa_bytes(raw)
    c=collections.Counter(); first={}
    for i in range(len(b)-15):
        if b[i]==0x1d and b[i+5]==0x25 and b[i+10]==0x2d:
            r,g,bl=[struct.unpack('<f',b[i+1+5*k:i+5+5*k])[0] for k in range(3)]
            if all(0<=v<=1 for v in (r,g,bl)):
                h='#%02X%02X%02X'%tuple(round(v*255) for v in (r,g,bl)); c[h]+=1; first.setdefault(h,i)
    print(name, dict(c.most_common(400)))
