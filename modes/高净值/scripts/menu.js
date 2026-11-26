function run(argv){
  const se=Application('System Events');
  const pr=se.processes.whose({bundleIdentifier:'com.apple.Keynote'})[0];
  const m=pr.menuBars[0].menus.byName(argv[0]);
  const it=m.menuItems.byName(argv[1]);
  const en=it.enabled(); if(en) it.click();
  return JSON.stringify({en});
}
