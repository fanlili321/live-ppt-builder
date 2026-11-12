#!/usr/bin/env python3
"""螺丝钉数据工具箱 开放接口 API v1 —— 只读取数，不需要浏览器登录。

凭证：环境变量 DBMS_API_TOKEN，或文件 ~/.config/dbms/token（权限 600）。
凭证由工具箱管理员在「API Client → 管理凭证」生成，默认有效期 90 天。
不要把凭证写进 skill、聊天、代码仓库。

用法：
  python3 dbms_api.py check
  python3 dbms_api.py star screw 2020-01-01 2026-11-06 --out screw.json   # 星级按区间取全（screw/hk/dividend/global/gold）
  python3 dbms_api.py get /funds/000001/nav beginDate=2026-01-01
  python3 dbms_api.py get /valuations/dingda indexCode=000688 beginDate=2026-08-01 endDate=2026-08-14
  python3 dbms_api.py index-query '{"codes":["000985.CSI"],...}'   # POST /index-queries（Wind 行情，只读查询）

只允许 GET 和 POST /index-queries；其他写接口一律拒绝。会自动翻页（meta 里有分页时）。
"""
import json
import os
import sys
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

BASE = 'https://dbms.yhlsd.com/api/v1'
TOKEN_FILE = Path.home() / '.config' / 'dbms' / 'token'


def token() -> str:
    t = os.environ.get('DBMS_API_TOKEN', '').strip()
    if not t and TOKEN_FILE.exists():
        if TOKEN_FILE.stat().st_mode & 0o077:
            os.chmod(TOKEN_FILE, 0o600)
        t = TOKEN_FILE.read_text().strip()
    if not t:
        sys.exit('没有工具箱凭证：请用户把凭证存到 ~/.config/dbms/token（见 rules/数据工具箱dbms.md）')
    return t


def call(method: str, path: str, params=None, body=None):
    url = BASE + path + ('?' + urllib.parse.urlencode(params) if params else '')
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(url, data=data, method=method, headers={
        'Authorization': 'Bearer ' + token(), 'Accept': 'application/json',
        **({'Content-Type': 'application/json'} if data else {})})
    try:
        with urllib.request.urlopen(req, timeout=60) as r:  # noqa
            return json.load(r)
    except urllib.error.HTTPError as e:
        msg = e.read().decode('utf-8', 'replace')[:300]
        if e.code in (401, 403):
            sys.exit(f'凭证无效、过期或无权限（HTTP {e.code}）：{msg}\n→ 请用户找管理员重新生成凭证并替换 ~/.config/dbms/token')
        sys.exit(f'HTTP {e.code}：{msg}')


def get_all(path: str, params: dict):
    """GET 并自动翻页，返回合并后的 data 列表（或单个对象）。"""
    params = dict(params)
    out, page = [], int(params.get('page', 1))
    while True:
        if page > 1:
            params['page'] = page
        res = call('GET', path, params)
        data, meta = res.get('data'), res.get('meta') or {}
        if not isinstance(data, list):
            return res
        out += data
        pg = meta.get('pagination') or meta
        total_pages = pg.get('totalPages') or pg.get('pages')
        if not total_pages or page >= int(total_pages):
            break
        page += 1
    return {'data': out, 'meta': {'count': len(out), 'path': path, 'params': params}}


def star_range(kind: str, begin: str, end: str) -> dict:
    """星级历史按区间取全：begin+end+limit=1000（单次上限 1000 条，按日期倒序），
    truncated 时把 end 前移到本批最早日期前一天继续取。gold 一次返回全部周度记录。
    注意：参数必须分开传（urlencode 自动用 & 连接）；2026-11-10 曾因测试时把多个参数
    拼成一个字符串传给 begin 导致服务端 500，非接口问题（服务端现已改为返回 422 校验错误）。"""
    import datetime as _dt
    out, seen, cur = [], set(), end
    for _ in range(50):
        d = call('GET', f'/star-ratings/{kind}', {'begin': begin, 'end': cur, 'limit': 1000})['data']
        items = d.get('items', [])
        if not items:
            break
        key = next(k for k in ('ratingDate', 'tradeDate', 'reportDate') if k in items[0])
        for x in items:
            if begin <= x[key] <= end and x[key] not in seen:
                seen.add(x[key]); out.append(x)
        if not d.get('truncated'):
            break
        oldest = min(x[key] for x in items)
        cur = (_dt.date.fromisoformat(oldest) - _dt.timedelta(days=1)).isoformat()
    key = next((k for k in ('ratingDate', 'tradeDate', 'reportDate') if out and k in out[0]), 'reportDate')
    out.sort(key=lambda x: x[key])
    return {'data': out, 'meta': {'count': len(out), 'kind': kind, 'begin': begin, 'end': end}}


def main():
    a = sys.argv[1:]
    if not a or a[0] not in ('check', 'get', 'index-query', 'star'):
        sys.exit(__doc__)
    out = None
    if '--out' in a:
        i = a.index('--out'); out = a[i + 1]; a = a[:i] + a[i + 2:]
    if a[0] == 'check':
        res = call('GET', '/star-ratings/screw', {'limit': 1})
        print('凭证可用；示例返回 meta：', json.dumps(res.get('meta'), ensure_ascii=False)[:300])
        return
    if a[0] == 'star':
        res = star_range(a[1], a[2], a[3])
    elif a[0] == 'get':
        params = dict(kv.split('=', 1) for kv in a[2:])
        res = get_all(a[1], params)
    else:
        res = call('POST', '/index-queries', body=json.loads(a[1]))
    text = json.dumps(res, ensure_ascii=False, indent=1)
    if out:
        Path(out).write_text(text)
        n = len(res['data']) if isinstance(res.get('data'), list) else 1
        print(f'已保存 {out}（{n} 条）')
    else:
        print(text)


if __name__ == '__main__':
    main()
