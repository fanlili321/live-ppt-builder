ObjC.import('CoreGraphics');
function click(x,y){
  const p=$.CGPointMake(x,y);
  $.CGEventPost($.kCGHIDEventTap,$.CGEventCreateMouseEvent(null,$.kCGEventMouseMoved,p,0)); delay(0.15);
  $.CGEventPost($.kCGHIDEventTap,$.CGEventCreateMouseEvent(null,$.kCGEventLeftMouseDown,p,0)); delay(0.05);
  $.CGEventPost($.kCGHIDEventTap,$.CGEventCreateMouseEvent(null,$.kCGEventLeftMouseUp,p,0)); delay(0.5);
}
function run(argv){
  const kn=Application('/Applications/Keynote Creator Studio.app'); kn.activate(); delay(0.8);
  const se=Application('System Events');
  const pr=se.processes.whose({bundleIdentifier:'com.apple.Keynote'})[0];
  click(+argv[0],+argv[1]);
  const m=pr.menuBars[0].menus.byName('排列');
  const names=m.menuItems.name();
  const it=m.menuItems.byName('移到最后面');
  const en=it.enabled();
  if(en) it.click();
  return JSON.stringify({names, en});
}
