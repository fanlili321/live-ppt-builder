import { spawnSync } from 'node:child_process';
import { writeFileSync } from 'node:fs';
import { join } from 'node:path';

const WIND_DIR = '/Users/fanlili/.agents/skills/wind-mcp-skill';
const OUT_DIR = '/Users/fanlili/Downloads/live-ppt-tool/current';
const items = [
  ['885189.WI', '451_p2_885189_WI.json'],
  ['SPX.GI', '451_p2_SPX_GI.json'],
  ['H00985.CSI', '451_p2_H00985_CSI.json'],
  ['HSI.HI', '451_p2_HSI_HI.json'],
];

for (const [windcode, filename] of items) {
  const params = JSON.stringify({
    windcode,
    begin_date: '20240902',
    end_date: '20251226',
    period: '11',
  });
  const r = spawnSync('node', ['scripts/cli.mjs', 'call', 'index_data', 'get_index_kline', params], {
    cwd: WIND_DIR,
    encoding: 'utf8',
    timeout: 90000,
  });
  if (r.status !== 0) {
    throw new Error(`${windcode} failed: ${r.stderr || r.stdout}`);
  }
  writeFileSync(join(OUT_DIR, filename), r.stdout);
  console.log(`saved ${windcode} -> ${filename}`);
}
