// osascript -l JavaScript gui_axis.js 文档 L T S 页 图表中心x 图表中心y max min steps
ObjC.import('CoreGraphics');
function click(x,y){const p=$.CGPointMake(x,y);
  $.CGEventPost($.kCGHIDEventTap,$.CGEventCreateMouseEvent(null,$.kCGEventMouseMoved,p,0)); delay(0.15);
  $.CGEventPost($.kCGHIDEventTap,$.CGEventCreateMouseEvent(null,$.kCGEventLeftMouseDown,p,0)); delay(0.05);
  $.CGEventPost($.kCGHIDEventTap,$.CGEventCreateMouseEvent(null,$.kCGEventLeftMouseUp,p,0)); delay(0.5);}
function run(argv){
  const [doc,L,T,S,sl,cx,cy,mx,mn,st]=[argv[0],+argv[1],+argv[2],+argv[3],+argv[4],+argv[5],+argv[6],argv[7],argv[8],argv[9]]; const dx=L-273, dy=T-328;
  const kn=Application('Keynote'); const d=kn.documents.byName(doc);
  const se=Application('System Events'); const pr=se.processes.byName('Keynote');
  d.currentSlide=d.slides[sl-1]; delay(1.0); kn.activate(); delay(0.5);
  click(L+1500*S,T+1065*S); click(L+cx*S,T+cy*S); delay(0.6);
  const w=pr.windows[0];
  const find=(role,pred)=>{for(const e of w.entireContents()){try{if(e.role()===role&&pred(e))return e}catch(x){}}return null};
  let tab=null; for(let k=0;k<5&&!tab;k++){ tab=find('AXRadioButton',e=>(e.description()||'')==='坐标轴'); if(!tab) delay(0.8);} if(!tab) return 'no axis tab';
  tab.actions.byName('AXPress').perform(); delay(0.8);
  const y=find('AXRadioButton',e=>(e.description()||'').indexOf('值')===0); if(y){ y.actions.byName('AXPress').perform(); delay(0.6);}
  const fields=[]; for(const e of w.entireContents()){try{if(e.role()==='AXTextField'){const p=e.position(); if(p[0]>1000+dx&&p[1]>290+dy&&p[1]<360+dy) fields.push([p[0],p[1],e])}}catch(x){}}
  const get=(x,yy)=>{for(const f of fields){if(Math.abs(f[0]-x)<15&&Math.abs(f[1]-yy)<10)return f[2]}return null};
  const fMax=get(1041+dx,312+dy), fMin=get(1041+dx,344+dy), fSt=get(1146+dx,312+dy);
  if(!fMax||!fMin||!fSt) return 'fields not found '+JSON.stringify(fields.map(f=>[f[0],f[1]]));
  const setf=(f,v)=>{f.focused=true; delay(0.2); f.value=v; delay(0.2); f.actions.byName('AXConfirm').perform(); delay(0.6);};
  setf(fMin,mn); setf(fMax,mx); setf(fSt,st);
  const r=[fMax.value(),fMin.value(),fSt.value()];
  click(L+1500*S,T+1065*S);
  return 'set '+JSON.stringify(r);
}
