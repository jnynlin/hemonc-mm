#!/usr/bin/env python3
"""組 mm_app.js：轉換核心取自 mm_convert.html，儀表板取自 mm_dashboard.html。"""
import json, re, sys, datetime
SP = sys.argv[1]
src = open(f'{SP}/mmapp/mm_app_src.js', encoding='utf8').read()
conv = open(f'{SP}/mmconv/mm_convert.html', encoding='utf8').read()
core = conv[conv.index('/*CORE-START*/'):conv.index('/*CORE-END*/') + len('/*CORE-END*/')]
dash = open(f'{SP}/mm_dashboard.html', encoding='utf8').read()
css = '\n'.join(re.findall(r'<style>([\s\S]*?)</style>', dash))
scripts = re.findall(r'<script>([\s\S]*?)</script>', dash)
assert len(scripts) == 1, len(scripts)
body = dash[dash.rindex('</style>') + len('</style>'):dash.index('<script>', dash.rindex('</style>'))].strip()
ver = sys.argv[2] if len(sys.argv) > 2 else datetime.date.today().isoformat()
qr = open(f'{SP}/mmapp/qr.js', encoding='utf8').read()
tx = dash[dash.index('/*TX-START*/'):dash.index('/*TX-END*/') + len('/*TX-END*/')]
out = (src.replace('__CORE__', core).replace('__TXRULES__', tx).replace('__QR__', qr)
          .replace('__DASH_CSS__', json.dumps(css, ensure_ascii=False))
          .replace('__DASH_BODY__', json.dumps(body, ensure_ascii=False))
          .replace('__DASH_JS__', json.dumps(scripts[0], ensure_ascii=False))
          .replace('__VERSION__', ver))
for tok in ('__TXRULES__', '__QR__', '__CORE__', '__DASH_CSS__', '__DASH_BODY__', '__DASH_JS__', '__VERSION__'):
    assert tok not in out, tok
open(f'{SP}/mmapp/mm_app.js', 'w', encoding='utf8').write(out)
print('mm_app.js', len(out.encode()), 'bytes, version', ver)
