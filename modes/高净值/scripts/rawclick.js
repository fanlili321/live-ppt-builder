ObjC.import('CoreGraphics');
function run(argv){
  const p=$.CGPointMake(+argv[0],+argv[1]);
  $.CGEventPost($.kCGHIDEventTap,$.CGEventCreateMouseEvent(null,$.kCGEventMouseMoved,p,0)); delay(0.3);
  $.CGEventPost($.kCGHIDEventTap,$.CGEventCreateMouseEvent(null,$.kCGEventLeftMouseDown,p,0)); delay(0.08);
  $.CGEventPost($.kCGHIDEventTap,$.CGEventCreateMouseEvent(null,$.kCGEventLeftMouseUp,p,0)); delay(0.5);
  return 'ok';
}
