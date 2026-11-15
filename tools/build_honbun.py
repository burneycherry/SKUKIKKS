#!/usr/bin/env python3
"""docs/honsho の読み取り記録（作成例）から、本文シート用のデータ data/honbun.js を作る。

  python3 tools/build_honbun.py          # data/honbun.js を書き出す
  python3 tools/build_honbun.py --text   # 変換結果を文字で表示（確認用）
  python3 tools/build_honbun.py --dropped  # 注記として除いた行の一覧（本文の行が消えていないかの確認用）

変換の決まり
- 「## 数字. 見出し」の節だけを使う（前書き・表紙・別添・読み取りの注意などは使わない）
- 箇条書きは字下げの深さつきの行、表は表として取り込む
- 記録用の注記（図の説明、照合の注記、ほかの工種との違いの説明など）は取り除く。図の行は「（図）」に置き換える
- 「(N) と同文」は、その工種の同じ番号の項から本文を持ってくる
"""
import glob
import json
import os
import re
import sys

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
DOCS = os.path.join(ROOT, 'docs', 'honsho')

# 作成例の番号 → SekokanExport.KOSHU の添字（(16) 総合試運転報告書は工種別施工計画書一覧に無い）
KOSHU = {1: 0, 2: 1, 3: 2, 4: 3, 5: 4, 6: 5, 7: 6, 8: 7, 9: 8, 10: 9, 11: 10, 12: 11, 13: 12, 14: 13, 15: 14, 17: 15}

DROP = re.compile(r'未照合|読み取り|原本|記録しない|`|利用者|と見られる|PDF p\.|アプリ|K-\d|は記録|だけ記録|出典|写真では|'
                  r'^（?続き|（p\.\d+ はここまで|ではない）$|と違[いう]|下の「|に記録$|^本文は「|'
                  r'^第Ⅲ編 3\.2 の作成例|の項[はが]無い|の図は、標準図|^見出しは「|は無い）?$|^列は |^枠\d|^表\d（|のとおりに読んだ|^形は|^備考（枠|太枠')
FIG = re.compile(r'^(図|図表|【参考図】|様式「|表「|「参考例)[（:「 ：]|^図$')
REF = re.compile(r'\((\d+)\)[^）]*?と同[文じ]|\((\d+)\)[^、。]*の [\d.、]+ と同文')


# 利用者の指示で、原本から表記を変えて出力するもの（00_kakunin.md の番号）
FIX = [
    ('工事着手前に打合せ', '工事着手前の打合せ'),                         # K-46
    ('EM-UTPケーブル（4対を超えるもの） | 仕上がり外径の10倍以上',            # K-42（表の行は「 | 」でつないで照合）
     'EM-UTPケーブル（4対を超えるもの） | 製造者説明書推奨値'),
    ('地表面下0.75m以下', '地表面下0.75m以上'),                             # K-55
    ('自家用工作物使用前', '自家用電気工作物使用前'),                         # K-72
    ('225以下 | 14mm2 以上', '250以下 | 14mm2 以上'),                        # K-38（標仕 表2.2.1 に合わせる）
]


def fix(t):
    for a, b in FIX:
        t = t.replace(a, b)
    return t


def clean(t):
    t = fix(t.replace('**', ''))
    # 行末に付けた記録用のかっこ書き（原本・アプリ・利用者の話）を取り除く
    t = re.sub(r'（[^（）]*(?:（[^（）]*）[^（）]*)*?(原本|アプリ|利用者|読み取り).*$', '', t)
    t = re.sub(r'（p\.[\d〜、・ 上下]+[^）]*）', '', t)
    t = re.sub(r'（\(\d+\)[^（）]*(（[^）]*）[^（）]*)*）', '', t)       # （(1) 屋内配管工事と同文）など
    t = re.sub(r'（[^（）]*と同文[^（）]*）', '', t)
    t = re.sub(r'（[^（）]*(読み取り|利用者|原本|p\.\d|と同じ)[^（）]*）', '', t)
    t = re.sub(r'（項目[:：][^（）]*(（[^）]*）[^（）]*)*）', '', t)
    return t.strip()


def parse(path):
    """節ごとに [(深さ, 文字) または (深さ, [[表の行]])] を返す。{番号: (見出し, 行)}"""
    secs = {}
    cur = None
    for raw in open(path, encoding='utf-8').read().split('\n'):
        m = re.match(r'^## (\d+)[.、] ?(.+)$', raw)
        if m:
            n = int(m.group(1))
            title = clean(m.group(2))
            title = re.sub(r'（続き.*$', '', title)
            if n in secs:
                cur = secs[n][1]
            else:
                cur = []
                secs[n] = (title, cur)
            continue
        if raw.startswith('#'):
            cur = None
            continue
        if cur is None or not raw.strip():
            continue
        ind = len(raw) - len(raw.lstrip(' '))
        s = raw.strip()
        if s.startswith('|'):
            if re.match(r'^\|[-| ]+\|$', s):
                continue
            cells = [c.strip() for c in s.strip('|').split('|')]
            if cur and isinstance(cur[-1][1], list) and cur[-1][2] == 'tb':
                cur[-1][1].append(cells)
            else:
                cur.append([ind // 2, [cells], 'tb'])
        elif s.startswith('**'):
            cur.append([ind // 2, s.replace('**', ''), 'cap'])
        elif s.startswith('- '):
            cur.append([ind // 2, s[2:], 'p'])
        elif cur and cur[-1][2] == 'p':
            cur[-1][1] += s                       # 折り返しの続き
        else:
            cur.append([ind // 2, s, 'p'])
    return secs


def subtree(lines, i):
    """i 行目とその子（字下げが深い行）の範囲の終わりを返す"""
    j = i + 1
    while j < len(lines) and lines[j][0] > lines[i][0]:
        j += 1
    return j


def find(lines, key):
    for i, l in enumerate(lines):
        if l[2] == 'p' and l[1].startswith(key):
            return i
    return -1


ALL = {}
DROPPED = []

# 別シートで出力している部分などの置き換え。{作成例の番号: {行の先頭: [差し替える子の行]}}
BETSU = '（別シート「%s」を参照）'
HANNYU = {
    '① 搬入計画書': ['搬入日時、搬入場所、機材と荷姿、運搬方法、搬入車両、作業体制（名簿含む）、安全対策、免許資格などは、'
                '機器搬入計画書の記載どおりとする', BETSU % '機器搬入計画書'],
    '② 搬入時の体制表': [BETSU % '搬入時の体制表'],
    '④ 揚重機の配置計画': ['クレーンの搬入路とクレーン配置計画を作成する。作業半径を確認する', BETSU % '搬入経路・揚重計画ほか'],
    '⑤ クレーンの選定': ['実際に使用するクレーンの能力を確認する', BETSU % '搬入経路・揚重計画ほか'],
    '⑥ 作業スケジュール表': ['作業当日の作業スケジュール表を作成する', BETSU % '搬入経路・揚重計画ほか'],
    '⑦ ワイヤロープの選定': ['ワイヤロープの選定、玉掛荷姿などは、次のとおり行う', BETSU % 'ワイヤロープの選定'],
}
OVERRIDE = {
    8: dict(HANNYU),
    9: dict(HANNYU, **{
        '⑧ 据付け施工要領': ['キュービクル式発電機の設置状況及びアンカーの施工要領、引抜き強度の確認等は、次のとおり行う',
                       BETSU % '据付け施工要領', BETSU % 'アンカーの施工要領'],
        '2つ目の項目': None,
    }),
    10: {'① 搬入計画書': HANNYU['① 搬入計画書'], '② 搬入時の体制表': HANNYU['② 搬入時の体制表']},
}
SHI = '（別シート「%s」にフローとチェックリストを出力）'
OVERRIDE[14] = {'⑥ 自主検査の方法及び判定基準': [BETSU % '①〜⑤-ハ の検査項目']}
OVERRIDE[15] = {k: [SHI % k[2:]] for k in ['① 停電試験1', '② 復電試験1', '① 火災発生試験2', '② 停電試験2',
                                            '③ 火災鎮火試験2', '④ 復電試験2']}
OVERRIDE[0] = {
    '7.1 品質管理体制': ['当該工事の品質管理体制を次に示す', BETSU % '品質管理体制'],
    '8.3 安全衛生管理組織': ['安全衛生管理組織を次に示す', BETSU % '安全衛生管理組織'],
    '8.4 安全衛生管理活動の計画表': ['安全衛生管理活動の計画を次表に示す', BETSU % '安全衛生活動計画表'],
    '8.7 事故時等の連絡網': ['事故時等の連絡先・体制図を次に示す', BETSU % '事故時等の連絡網'],
}
# 総合施工計画書で、章全体を別シートで出力しているもの
SOGO_CHAP = {
    2: ['当該工事の概要は、次による', BETSU % '工事概要'],
    3: ['当該工事の組織表、電気保安体制を次に示す', BETSU % '組織表', BETSU % '電気保安体制'],
    6: ['工程の管理は、全体工程表により行う', BETSU % '全体工程表'],
    10: ['総合仮設計画を次に示す', BETSU % '総合仮設計画'],
}
# 「(N) と同じ内容。ただし次の点が違う」と記録した項は、自動の参照では正しく作れないため、本文を明示的に持つ
KANRI_1 = ['監理技術者は、事前に主任技術者へ施工計画書を説明する',
           '作業前に、主任技術者が作業員へ作業手順を説明する。製品に添付されている施工要領がある場合は、併せて説明する',
           '施工中は、作業員による施工確認を行う']
KANRI_4B = ('一工程が完了したら、主任技術者が「電気設備工事の施工管理（令和8年版）」の施工管理表の「管理値・管理基準」の'
            '品質を確保しているかを確認する')
OVERRIDE[5] = {'6.2 管理方法': KANRI_1 + [KANRI_4B,
                                     '上記確認後、監理技術者と主任技術者で施工状況を「品質管理チェックシート」（別添1）により確認する']}
OVERRIDE[6] = {'6.2 管理方法': KANRI_1 + [KANRI_4B, '上記確認後、監理技術者と主任技術者で施工状況を「品質管理チェックシート」により確認する']}
OVERRIDE[10]['6.2 管理方法'] = [
    '監理技術者は、事前に主任技術者へ施工計画書を説明する',
    '作業前に、主任技術者が作業員へ作業手順を説明する。製品に添付されている施工要領を併せて説明する',
    '施工中は、作業員による施工の確認を行う',
    '主任技術者は添付の「機器搬入計画書」を含む本施工計画書の 5.「施工要領」による施工手順どおりに作業が行われているか確認する',
    '上記確認後、監理技術者と主任技術者で据付け状況を「品質管理チェックシート」（別添1）により確認する',
    '※受注者が作成しているチェックシートを添付する']
ADD = {9: {'自家発電装置の機材受入時は': ['機材は、設計図書に定める品質及び性能を有することの証明となる資料（試験成績書等）を、監督職員に提出する。'
                                'ただし、あらかじめ監督職員の承諾を受けた場合は、資料の提出を後日とする']}}


def resolve(num, depth=0):
    """作成例 num の節を、同文の参照を解決して返す"""
    secs = ALL[num]
    out = {}
    for n, (title, lines) in secs.items():
        res = []
        i = 0
        while i < len(lines):
            lv, text, kind = lines[i]
            m = REF.search(text) if kind == 'p' else None
            if m and depth < 4:
                ref = int(m.group(1) or m.group(2))
                key = re.match(r'^([\d.]+ |[①-⑳] )', text)
                other = resolve(ref, depth + 1) if ref in ALL and ref != num else {}
                src = other.get(n, ('', []))[1]
                if re.search(r'の [\d.、]+ と同文|とも \(\d+\)|、[\d.]+ [^ ]+とも', text) and src and lv == 0:
                    res.extend([l[:] for l in src])                 # 節ごと同文
                    i = subtree(lines, i)
                    continue
                if key and src:
                    k = find(src, key.group(1).strip() + ' ')
                    if k >= 0:
                        e = subtree(src, k)
                        base = src[k][0]
                        own = lines[i + 1:subtree(lines, i)]
                        res.append([lv, clean(src[k][1]), 'p'])
                        if not own:                                  # 自分の子が無ければ参照先の子を使う
                            for l in src[k + 1:e]:
                                res.append([l[0] - base + lv, l[1], l[2]])
                        i += 1
                        continue
            res.append([lv, text, kind])
            i += 1
        out[n] = (title, res)
    return out


def tidy(lines, num=0):
    over = OVERRIDE.get(num, {})
    add = ADD.get(num, {})
    out = []
    i = 0
    while i < len(lines):
        lv, text, kind = lines[i]
        if kind == 'tb':
            rows = [fix(' | '.join(clean(c) for c in r)).split(' | ') for r in text]
            out.append({'t': 'tb', 'lv': lv, 'rows': rows})
            i += 1
            continue
        t = clean(text)
        hit = [k for k in over if t.startswith(k)]
        if hit:
            if over[hit[0]] is not None:
                out.append({'t': 'p', 'lv': lv, 'text': hit[0]})
                for x in over[hit[0]]:
                    out.append({'t': 'p', 'lv': lv + 1, 'text': x})
            i = subtree(lines, i)
            continue
        m = re.match(r'^((?:[①-⑳]|\d+\.\d+) ?[^。:：]*?)。(?:図|表|様式|フロー図|列は)', t)
        if m and kind == 'p':                     # 図・様式の説明だけの項は、見出しと置き場所だけにする
            out.append({'t': 'p', 'lv': lv, 'text': m.group(1)})
            e = subtree(lines, i)
            tbs = [l for l in lines[i + 1:e] if l[2] == 'tb' or (l[2] == 'p' and l[1].startswith('※'))]
            if any(l[2] == 'tb' for l in tbs):     # 表があれば表を残す
                for l in tbs:
                    if l[2] == 'tb':
                        out.append({'t': 'tb', 'lv': lv + 1, 'rows': [[clean(c) for c in r] for r in l[1]]})
                    else:
                        out.append({'t': 'p', 'lv': lv + 1, 'text': clean(l[1])})
            else:
                out.append({'t': 'fig', 'lv': lv + 1, 'text': '（図・様式は Excel で作成してください）'})
            i = e
            continue
        if t.startswith('様式「据付け用アンカー'):
            out.append({'t': 'p', 'lv': lv, 'text': BETSU % 'アンカーの施工要領'})
            i = subtree(lines, i)
            continue
        hit = [k for k in over if t.startswith(k)]
        if hit:
            if over[hit[0]] is not None:
                out.append({'t': 'p', 'lv': lv, 'text': hit[0]})
                for x in over[hit[0]]:
                    out.append({'t': 'p', 'lv': lv + 1, 'text': x})
            i = subtree(lines, i)
            continue
        if FIG.search(t):
            out.append({'t': 'fig', 'lv': lv, 'text': '（図・様式は Excel で作成してください）'})
            i = subtree(lines, i)
            continue
        m = re.match(r'^((?:[①-⑳]|\d+\.\d+) [^。（]+)。', t)
        if m and DROP.search(t) and not DROP.search(m.group(1)):
            t = m.group(1)                        # 見出しのあとに続けた記録用の文だけを除く
        if not t or DROP.search(t):
            if t:
                DROPPED.append((num, t))
            i += 1
            continue
        out.append({'t': 'cap' if kind == 'cap' else 'p', 'lv': lv, 'text': t})
        for k in add:
            if t.startswith(k):
                for x in add[k]:
                    out.append({'t': 'p', 'lv': lv, 'text': x})
        i += 1
    # 図の行が続くときは1つにまとめる
    res = []
    for o in out:
        if o['t'] == 'fig' and res and res[-1]['t'] == 'fig':
            continue
        res.append(o)
    return res


def build():
    for f in sorted(glob.glob(os.path.join(DOCS, '03_3_2_*.md'))):
        num = int(re.search(r'03_3_2_(\d+)_', f).group(1))
        ALL[num] = parse(f)
    data = {'koshu': {}, 'sogo': []}
    for num, idx in KOSHU.items():
        secs = resolve(num)
        chap = []
        for n in sorted(secs):
            title, lines = secs[n]
            chap.append({'no': n, 'title': title, 'lines': tidy(lines, num)})
        data['koshu'][idx] = chap
    sogo = parse(os.path.join(DOCS, '03_2_2_sogo_rei.md'))
    for n in sorted(sogo):
        title, lines = sogo[n]
        if n in SOGO_CHAP:
            body = [{'t': 'p', 'lv': 0, 'text': x} for x in SOGO_CHAP[n]]
        else:
            body = tidy(lines, 0)
        if n == 7:                                  # 7.8 の表は別シート「工種別施工計画書一覧」と同じ
            body = [b for b in body if b['t'] != 'tb']
            k = [i for i, b in enumerate(body) if b.get('text', '').startswith('なお、施工計画書のうち')]
            if k:
                body.insert(k[0] + 1, {'t': 'p', 'lv': 1, 'text': BETSU % '工種別施工計画書一覧'})
        data['sogo'].append({'no': n, 'title': title, 'lines': body})
    return data


def as_text(chaps):
    out = []
    for c in chaps:
        out.append('%d. %s' % (c['no'], c['title']))
        for l in c['lines']:
            if l['t'] == 'tb':
                for r in l['rows']:
                    out.append('  ' * (l['lv'] + 1) + '| ' + ' | '.join(r))
            else:
                out.append('  ' * (l['lv'] + 1) + l['text'])
    return '\n'.join(out)


def main():
    data = build()
    if '--dropped' in sys.argv:                  # 注記として除いた行の一覧（本文の行が混じっていないか確認する）
        for n, t in DROPPED:
            print('(%s) %s' % (n, t))
        return
    if '--text' in sys.argv:
        for k in sorted(data['koshu']):
            print('===== 工種', k)
            print(as_text(data['koshu'][k]))
        print('===== 総合')
        print(as_text(data['sogo']))
        return
    # 記録用の言葉が本文に残っていないか確認する（残っていたら変換の決まりを直す）
    bad = re.compile(r'原本|アプリ|利用者|読み取り|同文|と同じ|Claude|`')
    left = [l.get('text', ' | '.join(sum(l.get('rows', []), []))) for ch in list(data['koshu'].values()) + [data['sogo']]
            for c in ch for l in c['lines']]
    left = [t for t in left if bad.search(t)]
    if left:
        print('記録用の言葉が本文に残っています（%d 行）:' % len(left))
        for t in left:
            print('  ' + t[:100])
        sys.exit(1)
    out = os.path.join(ROOT, 'data', 'honbun.js')
    head = ('/* 本文シートのデータ（自動生成。手で編集しない）\n'
            ' * python3 tools/build_honbun.py で docs/honsho の作成例の記録から作る。\n'
            ' * koshu: SekokanExport.KOSHU の添字ごとの章、sogo: 総合施工計画書の章。\n'
            ' * 章 = { no, title, lines }。行 = { t: p(文) / cap(表題) / fig(図の置き場所) / tb(表), lv: 字下げ, text または rows }\n'
            ' */\n')
    with open(out, 'w', encoding='utf-8') as fp:
        fp.write(head + 'window.HONBUN_DATA = ' + json.dumps(data, ensure_ascii=False, separators=(',', ':')) + ';\n')
    n = sum(len(c['lines']) for ch in data['koshu'].values() for c in ch)
    print('工種 %d、行 %d、総合 %d 行 → data/honbun.js' % (len(data['koshu']), n, sum(len(c['lines']) for c in data['sogo'])))


if __name__ == '__main__':
    main()
