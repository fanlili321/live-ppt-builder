// osascript -l JavaScript gui_align_left.js 文档名 L T S 页:x:y ...  (x,y 为文字框内一点的幻灯片坐标)
ObjC.import('CoreGraphics');
function click(x,y){
  const p=$.CGPointMake(x,y);
  $.CGEventPost($.kCGHIDEventTap,$.CGEventCreateMouseEvent(null,$.kCGEventMouseMoved,p,0)); delay(0.15);
  $.CGEventPost($.kCGHIDEventTap,$.CGEventCreateMouseEvent(null,$.kCGEventLeftMouseDown,p,0)); delay(0.05);
  $.CGEventPost($.kCGHIDEventTap,$.CGEventCreateMouseEvent(null,$.kCGEventLeftMouseUp,p,0)); delay(0.5);
}
function run(argv){
  const [doc,L,T,S]=[argv[0],+argv[1],+argv[2],+argv[3]];
  const kn=Application('Keynote'); const d=kn.documents.byName(doc);
  const se=Application('System Events');
  const toScr=(x,y)=>[L+x*S,T+y*S]; const log=[];
  for(const t of argv.slice(4)){
    const [sl,x,y]=t.split(':').map(Number);
    d.currentSlide=d.slides[sl-1]; delay(1.0); kn.activate(); delay(0.5);
    click(...toScr(1500,1065));
    click(...toScr(x,y));
    const pr=se.processes.byName('Keynote'); const mi=pr.menuBars[0].menus.byName('格式').menuItems.byName('文本').menus[0].menuItems.byName('左对齐'); if(!mi.enabled()) return JSON.stringify({ok:false,why:'左对齐不可用',sl,log}); mi.click(); delay(0.6);
    click(...toScr(1500,1065));
    log.push(sl);
  }
  return JSON.stringify({ok:true,log});
}
