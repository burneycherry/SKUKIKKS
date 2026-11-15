#!/usr/bin/env python3
"""docs/sekocheck/*.md（施工チェックシート 令和7年版の読み取り記録）→ data/sekocheck.js 変換

記録を直したら再実行する:  python3 tools/build_sekocheck.py
原本から表記を変えたもの（確認事項 SC-nn で決めたもの）は FIX_* に書く。
"""
import json
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, 'docs', 'sekocheck')
OUT = os.path.join(ROOT, 'data', 'sekocheck.js')

# 章: (id, 記録ファイル, 見出し, 画面・シート名に使う短い名前)
CHAPTERS = [
    ('ippan', 'c1_ippan.md', '1. 一般共通事項', '一般共通事項'),
    ('haikan', 'c2_1_haikan.md', '2. 電力設備、通信・情報設備工事　1) 配管工事', '配管工事'),
    ('haisen', 'c2_2_haisen.md', '2. 電力設備、通信・情報設備工事　2) 配線工事', '配線工事'),
    ('kiki', 'c2_3_kiki.md', '2. 電力設備、通信・情報設備工事　3) 機器取付け工事', '機器取付け工事'),
    ('setchi', 'c2_4_setchi.md', '2. 電力設備、通信・情報設備工事　4) 接地工事', '接地工事'),
    ('rai', 'c2_5_rai.md', '2. 電力設備、通信・情報設備工事　5) 雷保護設備工事', '雷保護設備工事'),
    ('kakuu', 'c2_6_kakuu.md', '2. 電力設備、通信・情報設備工事　6) 架空配線工事', '架空配線工事'),
    ('chichu', 'c2_7_chichu.md', '2. 電力設備、通信・情報設備工事　7) 地中配線工事', '地中配線工事'),
    ('juhen', 'c3_juhen.md', '3. 受変電設備・電力貯蔵設備・発電設備工事', '受変電・電力貯蔵・発電設備工事'),
    ('kanshi', 'c4_kanshi.md', '4. 中央監視制御設備工事', '中央監視制御設備工事'),
    ('iryo', 'c5_iryo.md', '5. 医療関係設備工事', '医療関係設備工事'),
]

# 本文の置換（SC-10）
FIX_TEXT = [
    ('脱落防止及びびつりボルトの要否', '脱落防止及びつりボルトの要否'),
]


def fix_ref(chap, kubun, ref):
    if ref == '1-1.1(5)':                       # SC-02 標仕 1.1.5 書面の書式及び取扱い
        return '1-1.1.5'
    if re.match(r'^6-2\.28\.2\((?!1\))', ref):  # SC-08・11 ⑴試験 の (1) を補う
        return ref.replace('6-2.28.2(', '6-2.28.2(1)(', 1)
    if chap == 'kakuu' and kubun == '接地' and ref == '2-2.1.6':   # SC-13 架空配線の接地
        return '2-2.11.6'
    return ref



def split_refs(text):
    """本文から 〔項番号〕 と直後の「他」を取り出す。戻り値 (本文, [項番号...])"""
    refs = []
    for m in re.finditer(r'〔([^〕]+)〕(他)?', text):
        refs.append(m.group(1))
        if m.group(2):
            refs.append('他')
    body = re.sub(r'〔[^〕]+〕(他)?', '', text)
    body = re.sub(r'\s+／', '／', body)
    body = re.sub(r'[ 　]+', ' ', body).strip()
    return body, refs


def clean(text):
    text = re.sub(r'【曖昧】（[^）]*）', '', text)
    for a, b in FIX_TEXT:
        text = text.replace(a, b)
    return text.strip()


def parse_chapter(cid, fname):
    sheets = []          # [{stage, kubun:[{name, mark, items:[{t, ref:[[行の番号, 項番号]]}]}]}]
    cur_sheet = None
    cur_kubun = None
    cur_item = None
    with open(os.path.join(SRC, fname), encoding='utf-8') as f:
        lines = f.read().split('\n')
    for ln in lines:
        m = re.match(r'^## p\.\d+(?:・\d+)? ?(.*)$', ln)
        if m:
            title = m.group(1)
            if not title:
                cur_kubun = None
                cur_item = None
                continue
            stage = '準備段階' if '準備段階' in title else ('施工段階' if '施工段階' in title else '')
            if cur_sheet is None or cur_sheet['stage'] != stage:
                cur_sheet = {'stage': stage, 'kubun': []}
                sheets.append(cur_sheet)
                cur_kubun = None
                cur_item = None
            continue
        m = re.match(r'^### (?:\[(検|立)\] )?(.+)$', ln)
        if m:
            mark, name = m.group(1) or '', m.group(2).strip()
            last = cur_sheet['kubun'][-1] if cur_sheet['kubun'] else None
            if last and last['name'] == name and last['mark'] == mark:
                cur_kubun = last        # 前のページから続く同じ区分（「（続き）」の項目は直前の項目にまとめる）
            else:
                cur_kubun = {'name': name, 'mark': mark, 'items': []}
                cur_sheet['kubun'].append(cur_kubun)
                cur_item = None         # 別の区分の項目に「（続き）」をまとめないようにする
            continue
        if cur_kubun is None:
            continue
        m = re.match(r'^- (.+)$', ln)
        if m:
            text = clean(m.group(1))
            if text.startswith('（以下'):
                continue
            body, refs = split_refs(text)
            refs = [fix_ref(cid, cur_kubun['name'], r) for r in refs]
            num = re.match(r'^(\d+\))', body)
            if body.endswith('（続き）') and cur_item is not None and num and \
                    cur_item['t'].startswith(num.group(1)):
                cur_item['ref'] += [[0, r] for r in refs]     # 続きのページの見出し行の番号は項目の見出し行に付ける
                continue
            # ref は [行の番号, 項番号]。行の番号は t を改行で分けたときの位置（0 = 項目の見出し行）
            cur_item = {'t': body, 'ref': [[0, r] for r in refs]}
            cur_kubun['items'].append(cur_item)
            continue
        m = re.match(r'^  - (.+)$', ln)
        if m and cur_item is not None:
            text = clean(m.group(1))
            body, refs = split_refs(text)
            refs = [fix_ref(cid, cur_kubun['name'], r) for r in refs]
            cur_item['t'] += '\n　' + body
            idx = cur_item['t'].count('\n')
            cur_item['ref'] += [[idx, r] for r in refs]
    # 項目の無い区分は出力できないため除き、知らせる
    for sh in sheets:
        empty = [k['name'] for k in sh['kubun'] if not k['items']]
        if empty:
            print('警告: 項目の無い区分を除きました', fname, sh['stage'], empty, file=sys.stderr)
        sh['kubun'] = [k for k in sh['kubun'] if k['items']]
    return sheets


def md_rows(text, ncols):
    rows = []
    for ln in text.split('\n'):
        if not ln.startswith('|') or ln.startswith('|---'):
            continue
        cells = [c.strip() for c in ln.strip().strip('|').split('|')]
        if len(cells) != ncols:
            continue
        rows.append(cells)
    return rows


def parse_shinsei():
    with open(os.path.join(SRC, 's1_shinsei.md'), encoding='utf-8') as f:
        s = f.read()
    main, beppyo = s.split('## p.114', 1)
    rows = [r for r in md_rows(main, 8) if r[0] != '大区分']
    prev = None
    out = []
    for r in rows:
        r = [c.replace('特性建設資材', '特定建設資材')            # SC-30
             .replace('特定建設業の種類', '特定建設作業の種類')    # SC-31
             for c in r]
        if prev:
            if r[1] == '同上':
                r[1] = prev[1]
            for i in range(len(r)):
                if r[i] == '（上と同じ欄）':                      # SC-29
                    r[i] = prev[i]
        r = [c.replace('／', '\n') for c in r]
        out.append(r)
        prev = r
    notes = []
    for ln in main.split('\n'):
        m = re.match(r'^注(\d)\. (.+)$', ln)
        if m:
            notes.append(m.group(1) + '. ' + m.group(2).replace('居出', '届出'))   # SC-28
    brows = [r for r in md_rows(beppyo, 2) if r[0] != '区分']
    brows = [[a, b.replace('／', '\n')] for a, b in brows]
    kind = re.search(r'工事の種類: (.+)', beppyo).group(1).strip()
    return {
        'title': '主な官公署その他への申請手続一覧表',
        'heads': ['工事区分', '', '申請・届出の名称', '提出者', '提出先', '提出時期', '摘要', '法令'],
        'rows': out,
        'notes': notes,
        'beppyo': {'title': '〈電事則〉別表第二（第62条、第65条関係）', 'kind': kind, 'rows': brows},
    }


def main():
    data = {'source': '一般社団法人 公共建築協会「電気設備工事施工チェックシート 令和7年版」', 'chapters': []}
    total = 0
    for cid, fname, title, short in CHAPTERS:
        sheets = parse_chapter(cid, fname)
        n = sum(len(k['items']) for sh in sheets for k in sh['kubun'])
        total += n
        print('%-8s %-24s シート%d 項目%3d' % (cid, short, len(sheets), n))
        data['chapters'].append({'id': cid, 'title': title, 'name': short, 'sheets': sheets})
    data['shinsei'] = parse_shinsei()
    print('項目計', total, '／申請手続', len(data['shinsei']['rows']), '行')
    js = ('/* 自動生成: tools/build_sekocheck.py （手で編集しない。元は docs/sekocheck/*.md） */\n'
          'window.SEKOCHECK_DATA = ' + json.dumps(data, ensure_ascii=False, separators=(',', ':')) + ';\n')
    with open(OUT, 'w', encoding='utf-8') as f:
        f.write(js)
    print('wrote', os.path.relpath(OUT, ROOT), len(js.encode('utf-8')), 'bytes')


if __name__ == '__main__':
    sys.exit(main())
