// 跨页复制形状（AppleScript 不能复制形状）：在 Keynote 界面里点选源形状 → 拷贝 → 到目标页点空白处 → 粘贴 N 次。
// 用法：osascript -l JavaScript gui_paste_shape.js 文档名 画布左 画布上 缩放 "源页:x:y" "目标页:个数" ...
//   画布左/上/缩放：幻灯片左上角在屏幕上的点坐标和缩放（截屏量，见 rules/样式规范_485定稿实测.md「三角」）
//   源 x,y：源形状中心的幻灯片坐标（1920x1080）。粘贴后新形状排在该页 shapes 末尾，再用 AppleScript 定位。
// 需辅助功能权限；运行时不要动鼠标键盘。每页校验粘贴数量，不符即停止。
ObjC.import('CoreGraphics');
function click(x,y){
  const p=$.CGPointMake(x,y);
  $.CGEventPost($.kCGHIDEventTap,$.CGEventCreateMouseEvent(null,$.kCGEventMouseMoved,p,0)); delay(0.15);
  $.CGEventPost($.kCGHIDEventTap,$.CGEventCreateMouseEvent(null,$.kCGEventLeftMouseDown,p,0)); delay(0.05);
  $.CGEventPost($.kCGHIDEventTap,$.CGEventCreateMouseEvent(null,$.kCGEventLeftMouseUp,p,0)); delay(0.4);
}
function run(argv){
  const [doc, L, T, S] = [argv[0], +argv[1], +argv[2], +argv[3]];
  const kn=Application('Keynote'); const d=kn.documents.byName(doc);
  const pr=Application('System Events').processes.byName('Keynote');
  const menu=(name)=>pr.menuBars[0].menus.byName('编辑').menuItems.byName(name);
  const toScr=(x,y)=>[L+x*S, T+y*S]; const log=[];
  kn.activate(); delay(0.8);
  const [srcSlide,sx,sy]=argv[4].split(':').map(Number);
  d.currentSlide=d.slides[srcSlide-1]; delay(1.0); kn.activate(); delay(0.5);
  click(...toScr(1500,1060)); click(...toScr(sx,sy));
  if(!menu('拷贝').enabled()) return JSON.stringify({ok:false,why:'拷贝不可用（没选中源形状）'});
  menu('拷贝').click(); delay(0.6);
  for(const t of argv.slice(5)){
    const [sl,n]=t.split(':').map(Number);
    const before=d.slides[sl-1].shapes.length;
    d.currentSlide=d.slides[sl-1]; delay(1.0); kn.activate(); delay(0.4);
    click(...toScr(1500,1060));
    for(let i=0;i<n;i++){ if(!menu('粘贴').enabled()) return JSON.stringify({ok:false,why:'粘贴不可用',sl}); menu('粘贴').click(); delay(0.7); }
    const after=d.slides[sl-1].shapes.length; log.push([sl,before,after]);
    if(after-before!==n) return JSON.stringify({ok:false,why:'粘贴数量不符',log});
  }
  return JSON.stringify({ok:true,log});
}
