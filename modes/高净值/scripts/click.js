ObjC.import('CoreGraphics');
function run(argv){
  const kn=Application('/Applications/Keynote Creator Studio.app'); kn.activate(); delay(0.6);
  const p=$.CGPointMake(+argv[0],+argv[1]);
  $.CGEventPost($.kCGHIDEventTap,$.CGEventCreateMouseEvent(null,$.kCGEventMouseMoved,p,0)); delay(0.15);
  $.CGEventPost($.kCGHIDEventTap,$.CGEventCreateMouseEvent(null,$.kCGEventLeftMouseDown,p,0)); delay(0.05);
  $.CGEventPost($.kCGHIDEventTap,$.CGEventCreateMouseEvent(null,$.kCGEventLeftMouseUp,p,0)); delay(0.6);
  return 'ok';
}
