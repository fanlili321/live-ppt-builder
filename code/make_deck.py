#!/usr/bin/env python3
"""
通用直播PPT管线入口。
用法：python3 make_deck.py spec_453.json
spec 是一个 JSON：{ out_path, template?, windcode?, img_dir?, pages:[...] }
每页见 ppt_engine.build_deck 支持的 type：
  cover / table / chart / chart_dual / image / three_cycle / summary
换一期只需复制一份 spec_xxx.json，改里面的内容即可。
"""
import sys, json, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ppt_engine import build_deck, TMPL_PATH, WINDCODE, IMG_DIR


def main():
    if len(sys.argv) < 2:
        print('用法: python3 make_deck.py <spec.json>')
        sys.exit(1)
    spec = json.load(open(sys.argv[1], encoding='utf-8'))
    print(f"=== 生成: {sys.argv[1]} → {spec['out_path']} ===")
    build_deck(
        spec['pages'],
        spec['out_path'],
        tmpl=spec.get('template', TMPL_PATH),
        windcode=spec.get('windcode', WINDCODE),
        img_dir=spec.get('img_dir', IMG_DIR),
    )


if __name__ == '__main__':
    main()
