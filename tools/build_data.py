#!/usr/bin/env python3
"""管理表(xlsx) → data/kanrihyo.js 変換スクリプト

使い方: python3 tools/build_data.py
入力 : source/R08_denki_kanrihyo_1.xlsx 〜 9.xlsx
出力 : data/kanrihyo.js  (window.KANRI_DATA = {...};)

列の対応 (元の管理表):
 A区分 B項目 C標仕 D標準図 E細目 F要点 G管理値・管理基準 H管理方法
 I(承/検/立) J作成資料 K作成写真 L提出資料 M提出写真 N掲載ページ
"""
import json
import os
import re
import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, 'source')
OUT = os.path.join(ROOT, 'data', 'kanrihyo.js')


def txt(v):
    """セル値を文字列化し、前後の空行・行末空白を除去する"""
    if v is None:
        return ''
    s = str(v).replace('\r\n', '\n').replace('\r', '\n')
    lines = [ln.rstrip() for ln in s.split('\n')]
    while lines and lines[-1].strip() == '':
        lines.pop()
    while lines and lines[0].strip() == '':
        lines.pop(0)
    return '\n'.join(lines)


def one_line(v):
    return re.sub(r'\s+', '', txt(v))


def mark(v):
    """提出欄の丸印 (○/〇 の表記ゆれを ○ に統一)"""
    s = one_line(v)
    return '○' if ('○' in s or '〇' in s) else s


def pages(v):
    """掲載ページ: 改行で並ぶ番号を重複なしで連結"""
    seen = []
    for p in re.split(r'[\s,、]+', txt(v)):
        if p and p not in seen:
            seen.append(p)
    return ','.join(seen)


def build_chapter(no):
    path = os.path.join(SRC, 'R08_denki_kanrihyo_%d.xlsx' % no)
    ws = openpyxl.load_workbook(path).worksheets[0]
    title = txt(ws['A1'].value)
    name = re.sub(r'^\S\s*', '', title).strip()
    # 印刷範囲があればその最終行まで、なければ最終行まで
    last = ws.max_row
    if ws.print_area:
        pa = ws.print_area if isinstance(ws.print_area, str) else ws.print_area[0]
        m = re.search(r'\$(\d+)$', pa)
        if m:
            last = int(m.group(1))
    items = []
    kubun = ''
    item = None
    group = None
    for r in range(4, last + 1):
        v = [ws.cell(r, c).value for c in range(1, 15)]
        a, b = one_line(v[0]), one_line(v[1])
        c, d, e = txt(v[2]), txt(v[3]), one_line(v[4])
        body = [txt(v[5]), txt(v[6]), txt(v[7]), one_line(v[8]), txt(v[9]),
                txt(v[10]), mark(v[11]), mark(v[12]), pages(v[13])]
        if not (a or b or c or d or e or any(body)):
            continue
        if a:
            kubun = a
        if b and (item is None or b != item['name'] or kubun != item['kubun']):
            item = {'id': '%d-%d' % (no, len(items)), 'kubun': kubun,
                    'name': b, 'groups': []}
            items.append(item)
            group = None
        if item is None:
            raise SystemExit('項目のない行: sheet %d row %d' % (no, r))
        new_group = group is None or c or e
        if (new_group and group is not None and one_line(c) == one_line(group['c'])
                and e == group['e']):
            new_group = False  # 改ページ用の見出し繰り返し
            if d and d not in group['d']:
                group['d'] = (group['d'] + '\n' + d).strip()
        if new_group:
            group = {'id': '%s-%d' % (item['id'], len(item['groups'])),
                     'c': c, 'd': d, 'e': e, 'rows': []}
            item['groups'].append(group)
        if any(body):
            group['rows'].append(body)
    for it in items:
        it['groups'] = [g for g in it['groups'] if g['rows'] or g['e']]
    return {'id': no, 'title': title, 'name': name, 'items': items}


def main():
    chapters = [build_chapter(n) for n in range(1, 10)]
    data = {
        'source': '国土交通省 電気設備工事 施工管理 管理表（令和8年版）',
        'cols': ['要点', '管理値・管理基準', '管理方法', '承検立', '作成資料',
                 '作成写真', '提出資料', '提出写真', '掲載ページ'],
        'chapters': chapters,
    }
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, 'w', encoding='utf-8') as f:
        f.write('/* 自動生成: tools/build_data.py （手で編集しない） */\n')
        f.write('window.KANRI_DATA = ')
        json.dump(data, f, ensure_ascii=False, separators=(',', ':'))
        f.write(';\n')
    for ch in chapters:
        ng = sum(len(i['groups']) for i in ch['items'])
        nr = sum(len(g['rows']) for i in ch['items'] for g in i['groups'])
        print('%s: 項目%d 細目%d 行%d' % (ch['title'], len(ch['items']), ng, nr))
    print('->', OUT, os.path.getsize(OUT), 'bytes')


if __name__ == '__main__':
    main()
