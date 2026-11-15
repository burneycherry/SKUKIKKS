#!/usr/bin/env python3
"""施工チェックシートの出力を検査する（tools/test_export.js の出力に対して実行する）

  python3 tools/check_sekocheck.py /tmp/after

- seko.xlsx: 項番号がデータと同じ件数だけ出ているか、小項目の丸数字が同じ行の小項目と合うか、
  Excel の上限（409pt）で切り詰めた行が無いか、全章のシートがあるか
- seko_sub.xlsx: 選んだ章（一般共通事項・配管工事）のシートだけがあるか
問題があれば内容を表示して終了コード 1 を返す。要 openpyxl
"""
import json
import os
import re
import sys

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CIRC = re.compile(r'[①-⑳㉑-㉕]')


def load_data():
    js = open(os.path.join(ROOT, 'data', 'sekocheck.js'), encoding='utf-8').read()
    return json.loads(js[js.index('{'):js.rindex(';')])


def check_sheets(wb):
    """チェックシート本体のシート（名前が数字で始まる）を検査し、(項番号の件数, 問題の一覧) を返す"""
    nref = 0
    bad = []
    for ws in [w for w in wb.worksheets if w.title[0].isdigit()]:
        for r in range(4, ws.max_row + 1):
            t = ws.cell(r, 2).value or ''
            refs = [x for x in (ws.cell(r, 3).value or '').split('\n') if x]
            nref += len(refs)
            marks = set(m for ln in t.split('\n')[1:] for m in CIRC.findall(ln[:2]))
            for x in refs:
                m = CIRC.match(x)
                if m and m.group(0) not in marks:
                    bad.append('%s %d行: 項番号 %s の丸数字がこの行の小項目に無い' % (ws.title, r, x))
            h = ws.row_dimensions[r].height
            if h is not None and h >= 409:
                bad.append('%s %d行: 409pt で切り詰めている' % (ws.title, r))
    return nref, bad


def main():
    if len(sys.argv) < 2:
        print('出力先のディレクトリを指定してください')
        return 1
    out = sys.argv[1]
    d = load_data()
    errs = []
    wb = openpyxl.load_workbook(os.path.join(out, 'seko.xlsx'))
    nref, bad = check_sheets(wb)
    errs += bad
    want = sum(len(it['ref']) for c in d['chapters'] for s in c['sheets'] for k in s['kubun'] for it in k['items'])
    if nref != want:
        errs.append('項番号の件数: データ %d／Excel %d' % (want, nref))
    nsheet = sum(len(c['sheets']) for c in d['chapters'])
    got = len([w for w in wb.worksheets if w.title[0].isdigit()])
    if got != nsheet:
        errs.append('チェックシートのシート数: データ %d／Excel %d' % (nsheet, got))
    for name in ('施工チェックシート 表紙', '使用にあたって', '申請手続一覧表', '電事則別表第二'):
        if name not in wb.sheetnames:
            errs.append('シートが無い: ' + name)
    sub = os.path.join(out, 'seko_sub.xlsx')
    if os.path.exists(sub):
        names = [w for w in openpyxl.load_workbook(sub).sheetnames if w[0].isdigit()]
        if not names or not all(n.startswith('1 ') or n.startswith('2-1 ') for n in names):
            errs.append('章の一部選択: 一般共通事項と配管工事以外のシートがある、または無い ' + str(names))
    print('項番号 %d件／チェックシート %dシート' % (nref, got))
    for e in errs:
        print('NG', e)
    print('OK' if not errs else '問題 %d件' % len(errs))
    return 1 if errs else 0


if __name__ == '__main__':
    sys.exit(main())
