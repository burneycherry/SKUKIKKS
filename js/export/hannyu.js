/* 施工計画書 Excel(xlsx) 出力：機器搬入計画書・搬入時の体制表・ワイヤロープの選定・据付け施工要領・アンカーの施工要領
 * 分割したファイルは index.html の順（core → … → main）に読み込む。部品は window.SKUKIKKSExport._ で共有する
 */
(function (X) {
  'use strict';

  var FONT = X.FONT, KOSHU = X.KOSHU, addBlank = X.addBlank, box = X.box, hline = X.hline,
    lineCount = X.lineCount, orMaru = X.orMaru, portraitSheet = X.portraitSheet, style = X.style,
    textLine = X.textLine, vline = X.vline, wareki = X.wareki;

  /* ---------- 機器搬入計画書・搬入時の体制表（作成例 (8)〜(10) の 5.1 ①②にならう） ---------- */

  // 搬入・据付け工事の工種名から機器名を取り出す（それ以外の工種は空文字）
  function hannyuKiki(koshu) {
    var name = koshu >= 0 && koshu < KOSHU.length ? KOSHU[koshu] : '';
    var m = /^(.+)搬入・据付け工事施工計画書$/.exec(name);
    return m ? m[1] : '';
  }

  // 共通で決まる部分だけ記入し、機器ごとに変わる部分は ○○ のまま出力する（Excel で仕上げる）
  function addHannyu(wb, info, koshu) {
    var ws = portraitSheet(wb, '機器搬入計画書');
    var WV = 25;                        // 記事の1列分の幅（3列で1行）
    ws.columns = [{ width: 4 }, { width: 19 }, { width: WV }, { width: WV }, { width: WV }];
    var H = { size: 10, v: 'middle', h: 'center', fill: true };
    var K = { size: 9, v: 'middle', h: 'center' };
    var V = { size: 10, v: 'middle' };
    var kiki = hannyuKiki(koshu);
    var cub = koshu === 7;              // キュービクル式配電盤、直流電源装置（作成例の記載を使う）
    var gen = koshu === 8;              // 自家発電装置（作成例 (9) p.446 の記載を使う）
    var pv = koshu === 9;               // 太陽光発電装置（作成例 (10) p.461 の記載を使う）
    var juchu = info.juchu || '○○○○電気株式会社';

    // 機器・現場ごとに変わる欄（㈣〜㈥の記事）は白紙にして、項目だけ残す（利用者の指示）
    var hinmoku = [
      ['品名', gen ? 'キュービクル式自家発電装置' : (pv ? '太陽光発電装置' : (kiki ? kiki + 'ほか' : ''))],
      ['個数', '']
    ];
    if (gen) {
      hinmoku.push(['自家発電装置', '']);
    } else if (!pv) {
      hinmoku.push([cub ? '①〔キュービクル〕' : '①', '']);
      hinmoku.push([cub ? '②〔直流電源装置〕' : '②', '']);
    }
    var mokuteki = gen ? '「自家発電装置」の搬入据付工事'
      : (pv ? '「太陽光発電装置」の搬入・据付け工事'
        : '「電気設備機器（' + (kiki || '○○○○') + '）」の搬入据付工事');
    var keiro = pv ? ['(1) ', '(2) '] : (gen ? ['(1) ', '(2) ', '(3) '] : ['(1) ', '(2) ', '(3) ', '(4) ', '(5) ']);
    var nichiji = [
      ['開始', ['令和○○年○○月○○日（○）午前　○○時～']],
      ['終了', ['令和○○年○○月○○日（○）午後　○○時']]
    ];
    if (!pv) { nichiji.push(['特記', ['作業時間帯：○○：○○～○○：○○']]); }
    var sharyo = pv ? ['(1) 機器輸送車両', '(2) 機材車両（小型・手工具輸送）']
      : ['(1) 機器輸送車両', '(2) 機材車両（大型機材輸送）', '(3) 機材車両（小型・手工具輸送）'];
    var shisetsu = [
      '(1) 作業区域は「カラーコーン・バリケード」を設置して立入禁止とする。',
      '(2) ' + (pv ? '機材' : '機械') + '、運搬具等は必ず仮つりを行った後、本つりを行う。',
      '(3) ワイヤロープほかの使用機材の整備・点検を徹底する。'
    ];
    if (!pv) {
      shisetsu.push('(4) 作業照度不足の通路、' + (cub || gen ? '電気室' : '搬入場所') +
        'には投光器を設置し照度を確保し作業の安全に努める。');
    }

    // [区分, 項目, 記事...]。記事は1行ずつ。3要素の配列は3列に分けて出力する
    var G = [
      ['㈠', [['目的', [mokuteki]]]],
      ['㈡', [['搬入場所', ['○○建物：○○○室']]]],
      ['㈢実施日時', nichiji],
      ['㈣品目', hinmoku.map(function (h) { return [h[0], [h[1]]]; })],
      ['㈤搬入業者', [
        ['住所・電話', ['']],
        ['業者名', ['']],
        ['連絡者', ['']]
      ]],
      ['㈥搬入方法', [
        ['運搬方法\n及び経路', keiro],
        ['車両', sharyo],
        ['体制', ['(1) 組織表　　　　別紙（搬入時の体制表）', '(2) 作業員名簿　　別紙']]
      ]],
      ['㈦安全対策', [
        ['施設', shisetsu],
        ['人員', ['監視員を配置する。']],
        ['教育', [
          '(1) 各種役割の責任者及び新規入場者教育を受けた作業員を配置する。',
          '(2) 場内車両運行速度制限の厳守と第三者災害の防止を徹底する。'
        ]],
        ['緊急連絡先', [
          '(1) 施工者　　　' + juchu + '　　TEL ○○○○○○',
          '(2) 搬入業者　　○○○○株式会社　　TEL ○○○○○○'
        ]]
      ]],
      ['㈧資格他', [
        ['免許資格の\n種類\n（該当のもの\nに○印）', [
          ['クレーン・デリック運転士免許', (pv ? '小型' : '') + '移動式クレーン運転士技能講習', '建設用リフト特別教育修了者'],
          ['巻上げ機特別教育修了者', '玉掛け技能講習修了者', '足場の組立て等作業主任者'],
          ['高所作業車特別教育修了者', '高所作業車技能講習修了者', 'フォークリフトの運転業務']
        ]]
      ]],
      ['㈨その他', [
        ['申請書類', [
          '(1) 道路使用許可証　　許可番号（　　　　　　　　　　）',
          '(2) 通行許可証　　　　　〃　　（　　　　　　　　　　）',
          '(3) 特殊車両通行許可申請　　要　・　不要'
        ]],
        ['近隣対策', ['(1) 時間　　：早朝、昼間、夜間', '(2) 近隣対策：済み、未、無']]
      ]]
    ];

    textLine(ws, 1, 5, '機器搬入計画書', 16, true, 30);
    textLine(ws, 2, 5, '工事名　' + (info.kojiName || ''), 10, false, 20, 'left');
    box(ws, 3, 1, 3, 2, '項目', H);
    box(ws, 3, 3, 3, 5, '記事', H);
    ws.getRow(3).height = 20;

    var r = 4;
    G.forEach(function (g) {
      var g1 = r;
      g[1].forEach(function (it) {
        var k1 = r;
        it[1].forEach(function (line) {
          var n = 1;
          if (typeof line === 'string') {
            box(ws, r, 3, r, 5, line, V);
            n = lineCount(line, WV * 3, 10);
          } else {
            line.forEach(function (t, i) {
              box(ws, r, 3 + i, r, 3 + i, t, K);
              n = Math.max(n, lineCount(t, WV, 9));
            });
          }
          ws.getRow(r).height = Math.max(19, n * 13.5 + 6);
          r++;
        });
        box(ws, k1, 2, r - 1, 2, it[0], K);
      });
      box(ws, g1, 1, r - 1, 1, g[0], { size: 9, v: 'middle', h: 'center', vertical: true, wrap: false });
    });
    ws.getRow(r).height = 8;
    textLine(ws, r + 1, 5, '（○○ や空欄の箇所は、搬入する機器・現場・搬入業者に合わせて記入してください。' +
      '該当する免許資格・選択肢には ○ を付けてください。）', 9, false, 30, 'left');
    return ws;
  }

  function addHannyuTaisei(wb, info, koshu) {
    var ws = portraitSheet(wb, '搬入時の体制表');
    var cols = [];
    var i;
    for (i = 0; i < 14; i++) { cols.push({ width: 6.2 }); }
    ws.columns = cols;
    var LAST = 14;
    var B = { size: 11, v: 'middle', h: 'center' };
    var S = { size: 10, v: 'middle', h: 'center' };
    var koji = info.kojiName || '○○○○○○工事';
    var juchu = info.juchu || '○○○○電気工事㈱';
    var r;

    function tel(row) {
      var names = ['（会社）', '（現場）', '（携帯）'];
      names.forEach(function (n, k) {
        ws.mergeCells(row + k, 11, row + k, 14);
        var c = ws.getCell(row + k, 11);
        c.value = ' TEL ○○○○○○' + n;
        c.font = { name: FONT, size: 10 };
        c.alignment = { vertical: 'middle' };
      });
    }
    function person(row, role, name) {
      box(ws, row, 3, row, 10, koji, B);
      box(ws, row + 1, 3, row + 1, 10, juchu, B);
      box(ws, row + 2, 3, row + 2, 6, role, B);
      box(ws, row + 2, 7, row + 2, 10, name, B);
    }
    function team(row, heads, names) {
      heads.forEach(function (h, k) {
        var c1 = 1 + k * 2;
        box(ws, row, c1, row, c1 + 1, h, { size: 9, v: 'middle', h: 'center', fill: true });
        names.forEach(function (line, q) {
          box(ws, row + 1 + q, c1, row + 1 + q, c1 + 1, line[k] ? '○○○○' : '', S);
        });
        box(ws, row + 1 + names.length, c1, row + 1 + names.length, c1 + 1, '○○○○㈱', S);
      });
    }

    textLine(ws, 1, LAST, '搬入時の体制表', 16, true, 30);
    ws.getRow(2).height = 10;
    person(3, '現場代理人', orMaru(info.dairinin, '○○○○'));
    tel(3);
    vline(ws, 6, 6, 6);
    person(7, '監理(主任)技術者', orMaru(info.gijutsu, '○○○○'));
    vline(ws, 10, 10, 6);
    person(11, '現場担当', '○○○○');
    vline(ws, 14, 14, 6);
    box(ws, 15, 3, 15, 6, '搬入業者', B);
    box(ws, 15, 7, 15, 10, '○○○○㈱', B);
    box(ws, 16, 3, 16, 6, '現場責任者', B);
    box(ws, 16, 7, 16, 10, '○○○○', B);
    box(ws, 17, 3, 17, 6, '安全衛生責任者', B);
    box(ws, 17, 7, 17, 10, '○○○○', B);
    tel(15);
    vline(ws, 18, 19, 6);
    if (koshu === 9) {                  // 太陽光発電装置（作成例 (10) p.462）は作業班の表が1つ
      team(20, ['作業指揮者', 'クレーン\nオペレーター', '玉掛者', '合図者', '据付作業員', '監視員', '誘導員'],
        [[1, 1, 1, 1, 1, 1, 1], [0, 0, 0, 0, 1, 1, 0], [0, 0, 0, 0, 1, 0, 0]]);
    } else {
      hline(ws, 19, 7, 12);             // 2つ目の作業班へ回す線
      vline(ws, 19, 27, 12);
      hline(ws, 28, 9, 12);
      team(20, ['作業指揮者', 'クレーン\nオペレーター', '玉掛者', '合図者', '据付作業員'],
        [[1, 1, 1, 0, 1], [0, 1, 1, 0, 1], [0, 0, 0, 0, 1]]);
      team(26, ['据付作業員', '組立作業員', '監視員', '誘導員'],
        [[1, 1, 1, 1], [1, 1, 1, 0], [0, 0, 0, 0]]);
    }
    for (r = 3; r <= 30; r++) { ws.getRow(r).height = 22; }
    ws.getRow(6).height = 12; ws.getRow(10).height = 12; ws.getRow(14).height = 12;
    ws.getRow(18).height = 12; ws.getRow(19).height = 12; ws.getRow(25).height = 12;
    ws.getRow(20).height = 28; ws.getRow(26).height = 28;
    ws.getRow(31).height = 10;
    textLine(ws, 32, LAST, '（○○○○ の箇所は、氏名・会社名・電話番号に書き換えてください。枠はセルの罫線です。）',
      9, false, 18, 'left');
    return ws;
  }

  /* 搬入計画の図・表（作成例 (8) の 5.1 ③〜⑥）。表題だけ出力し、利用者が Excel で仕上げる */
  function addHannyuZu(wb, info, koshu) {
    var ws = addBlank(wb, info, '搬入経路・揚重計画ほか',
      '（このシートに ③〜⑥ の図・表を作成してください。見出しの位置は自由に動かせます。）');
    var items = [
      ['③ 搬入経路', (koshu === 8 ? '搬入ルートは' : '搬入は') + '、次の経路で行う。なお、搬入経路の壁面は、ベニヤ板で養生する。'],
      ['④ 揚重機の配置計画', 'クレーンの搬入路とクレーン配置計画を作成する。作業半径を確認する。'],
      ['⑤ クレーンの選定', '実際に使用するクレーンの能力を確認する。（定格総荷重表、作業半径－揚程図を添付）'],
      ['⑥ 作業スケジュール表', '作業当日の作業スケジュール表を作成する。']
    ];
    items.forEach(function (it, i) {
      var r = 5 + i * 12;
      textLine(ws, r, 20, it[0], 12, true, 22, 'left');
      textLine(ws, r + 1, 20, '・' + it[1], 10, false, 20, 'left');
    });
    return ws;
  }

  /* ワイヤロープの選定（作成例 (8) の 5.1 ⑦）。条件の欄を書き換えると Excel 上で再計算される */
  var WIRE_T1 = [[0, 1.00], [10, 1.005], [20, 1.02], [30, 1.04], [40, 1.07], [50, 1.10], [60, 1.16], [70, 1.22],
    [80, 1.31], [90, 1.41], [100, 1.56], [110, 1.74], [120, 2.00], [130, 2.37], [140, 2.93], [150, 3.86]];
  // [ロープの径, 4号めっき種, 4号A種(裸), 6号めっき種, 6号A種(裸), 6×24の標準重量]
  var WIRE_T2 = [
    [8, 2.97, 3.21, 3.19, 3.46, 0.212], [9, 3.75, 4.06, 4.04, 4.38, 0.269], [10, 4.64, 5.02, 4.99, 5.41, 0.332],
    [11.2, 5.82, 6.29, 6.26, 6.79, 0.416], [12.5, 7.25, 7.84, 7.80, 8.45, 0.519], [14, 9.09, 9.83, 9.81, 10.60, 0.651],
    [16, 11.90, 12.80, 12.80, 13.80, 0.850], [18, 15.00, 16.20, 16.20, 17.50, 1.080],
    [20, 18.50, 20.10, 19.90, 21.60, 1.330], [22.4, 23.30, 25.20, 25.00, 27.10, 1.670],
    [25, 29.00, 31.30, 31.20, 33.80, 2.080]
  ];

  function addWire(wb, info, koshu) {
    var ws = portraitSheet(wb, 'ワイヤロープの選定');
    ws.columns = [{ width: 13 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 },
      { width: 12 }, { width: 8 }];
    var LAST = 8;
    var H = { size: 10, v: 'middle', h: 'center', fill: true };
    var C = { size: 10, v: 'middle', h: 'center' };
    var L = { size: 10, v: 'middle' };
    var INPUT = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
    // 条件の初期値（作成例の値）
    var W = koshu === 8 ? 3.5 : 1.8, DIA = 16, N = 4, TH = 60, SF = 6, COL = 2;
    var coef = 1.16, cut = 12.80;
    function safe(c) { return Math.round(c * N / (SF * coef) * 1000) / 1000; }
    function fx(cell, formula, result) { cell.value = { formula: formula, result: result }; }

    textLine(ws, 1, LAST, 'ワイヤロープの選定', 16, true, 30);
    textLine(ws, 2, LAST, '工事名　' + (info.kojiName || ''), 10, false, 20, 'left');
    textLine(ws, 3, LAST, '1. 選定条件（色付きの欄を書き換えると、計算結果と表2の判定が再計算されます）', 11, true, 22, 'left');

    var conds = [
      ['つり荷（最大重量品）', koshu === 7 ? 'キュービクル式配電盤（最大重量品 ○○○○）'
        : (koshu === 8 ? 'キュービクル式発電機' : '○○○○'), ''],
      ['最大荷重 (t)', W, ''],
      ['ワイヤロープの径 (mm)', DIA, '表2の径で入力'],
      ['つり本数 (本)', N, ''],
      ['つり角度 θ (°)', TH, '表1の角度で入力（60°以内）'],
      ['安全係数', SF, 'クレーン等安全規則第213条（6以上）'],
      ['切断荷重の列（表2）', COL, '1:4号めっき 2:4号A種 3:6号めっき 4:6号A種']
    ];
    conds.forEach(function (c, i) {
      var r = 4 + i;
      box(ws, r, 1, r, 3, c[0], L);
      if (i === 0) {
        box(ws, r, 4, r, 8, c[1], L);
      } else {
        box(ws, r, 4, r, 4, c[1], C);
        box(ws, r, 5, r, 8, c[2], { size: 9, v: 'middle' });
      }
      ws.getCell(r, 4).fill = INPUT;
      ws.getRow(r).height = 20;
    });

    textLine(ws, 11, LAST, '2. 計算結果　安全荷重＝（切断荷重×つり本数）÷（安全係数×張力増加係数）', 11, true, 22, 'left');
    box(ws, 12, 1, 12, 3, '張力増加係数（表1）', L);
    box(ws, 12, 4, 12, 4, '', C);
    fx(ws.getCell(12, 4), 'IF(D8<=70,VLOOKUP(D8,D20:E27,2,FALSE),VLOOKUP(D8,F20:G27,2,FALSE))', coef);
    box(ws, 13, 1, 13, 3, '切断荷重（表2） (t)', L);
    box(ws, 13, 4, 13, 4, '', C);
    fx(ws.getCell(13, 4), 'INDEX(B34:E44,MATCH(D6,A34:A44,0),D10)', cut);
    box(ws, 14, 1, 14, 3, '安全荷重 (t)', L);
    box(ws, 14, 4, 14, 4, '', C);
    fx(ws.getCell(14, 4), 'ROUND(D13*D7/(D9*D12),3)', safe(cut));
    ws.getCell(14, 4).numFmt = '0.000';
    box(ws, 15, 1, 15, 3, '判定', L);
    box(ws, 15, 4, 15, 8, '', L);
    fx(ws.getCell(15, 4), 'IF(D14>D5,"安全荷重 ＞ 最大荷重　問題なし","安全荷重が不足　径・本数を見直す")',
      '安全荷重 ＞ 最大荷重　問題なし');
    box(ws, 16, 1, 16, 3, '使用できる最小の径 (mm)', L);
    box(ws, 16, 4, 16, 4, '', C);
    var minDia = '該当なし';
    WIRE_T2.forEach(function (t) {
      if (minDia === '該当なし' && safe(t[COL]) > W && t[0] >= 12.5) { minDia = t[0]; }
    });
    fx(ws.getCell(16, 4), 'IFERROR(INDEX(A34:A44,MATCH("○",H34:H44,0)),"該当なし")', minDia);
    box(ws, 16, 5, 16, 8, '表2の判定が ○ になる最小の径（安全荷重 ＞ 最大荷重、かつ 12.5mm 以上）', { size: 9, v: 'middle' });
    var r;
    for (r = 12; r <= 16; r++) { ws.getRow(r).height = 20; }
    ws.getRow(17).height = 8;

    // 図1 つり角度（左）と 表1（右）
    ws.mergeCells(18, 4, 18, 7);
    ws.getCell(18, 4).value = '表1　ワイヤロープ張力増加係数';
    ws.getCell(18, 4).font = { name: FONT, size: 10, bold: true };
    ws.getCell(18, 4).alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(18).height = 20;
    ['つり角度\n（θ°）', '張力増加\n係数', 'つり角度\n（θ°）', '張力増加\n係数'].forEach(function (t, i) {
      box(ws, 19, 4 + i, 19, 4 + i, t, H);
    });
    ws.getRow(19).height = 30;
    for (r = 0; r < 8; r++) {
      box(ws, 20 + r, 4, 20 + r, 4, WIRE_T1[r][0], C);
      box(ws, 20 + r, 5, 20 + r, 5, WIRE_T1[r][1], C);
      box(ws, 20 + r, 6, 20 + r, 6, WIRE_T1[r + 8][0], C);
      box(ws, 20 + r, 7, 20 + r, 7, WIRE_T1[r + 8][1], C);
      ws.getCell(20 + r, 5).numFmt = WIRE_T1[r][1] === 1.005 ? '0.000' : '0.00';
      ws.getCell(20 + r, 7).numFmt = '0.00';
      ws.getRow(20 + r).height = 17;
    }
    // 図: つり線（縦線）、三角形（斜めの罫線）、つり荷（枠）
    vline(ws, 19, 19, 1);
    ws.mergeCells(20, 1, 23, 1);
    ws.mergeCells(20, 2, 23, 2);
    ws.getCell(20, 1).border = { diagonal: { up: true, down: false, style: 'thin', color: { argb: 'FF000000' } } };
    ws.getCell(20, 2).border = { diagonal: { up: false, down: true, style: 'thin', color: { argb: 'FF000000' } } };
    ws.getCell(20, 3).value = '← つり角度 θ';
    ws.getCell(20, 3).font = { name: FONT, size: 9 };
    ws.getCell(22, 3).value = '← 玉掛ワイヤ';
    ws.getCell(22, 3).font = { name: FONT, size: 9 };
    box(ws, 24, 1, 26, 2, 'つり荷', { size: 10, v: 'middle', h: 'center' });
    ws.getCell(27, 1).value = '図1　つり角度';
    ws.getCell(27, 1).font = { name: FONT, size: 10, bold: true };
    textLine(ws, 28, LAST, '（注）張力増加係数＝1／cos（θ／2）。θ＝0（2本の玉掛用ワイヤロープが平行）のときのワイヤロープ張力を1とする。',
      9, false, 18, 'left');
    ws.getRow(29).height = 8;

    // 表2
    textLine(ws, 30, LAST, '表2　ワイヤロープの切断荷重 (t)　と　選定条件での安全荷重', 10, true, 20, 'left');
    box(ws, 31, 1, 33, 1, 'ロープの径\n(mm)', H);
    box(ws, 31, 2, 31, 3, '4号　6×24', H);
    box(ws, 31, 4, 31, 5, '6号　6×37', H);
    box(ws, 32, 2, 32, 2, '普通Z', H);
    box(ws, 32, 3, 32, 3, '普通Z又はS', H);
    box(ws, 32, 4, 32, 4, '普通Z', H);
    box(ws, 32, 5, 32, 5, '普通Z又はS', H);
    box(ws, 33, 2, 33, 2, 'めっき種', H);
    box(ws, 33, 3, 33, 3, 'A種（裸）', H);
    box(ws, 33, 4, 33, 4, 'めっき種', H);
    box(ws, 33, 5, 33, 5, 'A種（裸）', H);
    box(ws, 31, 6, 33, 6, '6×24の\n標準重量\n(kg/m)', H);
    box(ws, 31, 7, 33, 7, '安全荷重\n(t)', H);
    box(ws, 31, 8, 33, 8, '判定', H);       // 安全荷重 ＞ 最大荷重 かつ 径 12.5mm 以上で ○
    for (r = 31; r <= 33; r++) { ws.getRow(r).height = 17; }
    WIRE_T2.forEach(function (t, i) {
      var rr = 34 + i;
      var k;
      for (k = 0; k < 6; k++) {
        box(ws, rr, 1 + k, rr, 1 + k, t[k], C);
        if (k > 0) { ws.getCell(rr, 1 + k).numFmt = k === 5 ? '0.000' : '0.00'; }
      }
      box(ws, rr, 7, rr, 7, '', C);
      fx(ws.getCell(rr, 7), 'ROUND(INDEX(B' + rr + ':E' + rr + ',$D$10)*$D$7/($D$9*$D$12),3)', safe(t[COL]));
      ws.getCell(rr, 7).numFmt = '0.000';
      box(ws, rr, 8, rr, 8, '', C);
      fx(ws.getCell(rr, 8), 'IF(AND(G' + rr + '>$D$5,A' + rr + '>=12.5),"○","×")',
        (safe(t[COL]) > W && t[0] >= 12.5) ? '○' : '×');
      ws.getRow(rr).height = 17;
    });
    ws.getRow(45).height = 8;
    [
      '・切断荷重は JIS に定められた値とする。ワイヤロープのつり角度は 60°以内で使用する。',
      '・玉掛用具（ワイヤロープ）の安全係数は、クレーン等安全規則第213条により 6 以上とする。',
      '・ワイヤロープの径は、12.5mm（安全作業基準の一例）以上を使用すること。'
    ].forEach(function (t, i) { textLine(ws, 46 + i, LAST, t, 10, false, 18, 'left'); });
    return ws;
  }

  /* 据付け施工要領（作成例 (8) の 5.1 ⑧⑨ の本文だけ。図は利用者が Excel で作成する） */
  function addSuetsuke(wb, info, koshu) {
    if (koshu === 9) { return addSuetsukePv(wb, info); }
    var ws = portraitSheet(wb, '据付け施工要領');
    ws.columns = [{ width: 90 }];
    var kiki = hannyuKiki(koshu);
    var cub = koshu === 7;
    var name = cub ? 'キュービクル' : '機器';
    var r = 1;
    function line(t, size, bold, h) { textLine(ws, r, 1, t, size, bold, h, 'left'); r++; }
    textLine(ws, r, 1, '据付け施工要領', 16, true, 30); r++;
    line('工事名　' + (info.kojiName || ''), 10, false, 20);
    line('⑧ 据付け施工要領', 12, true, 24);
    line('　' + (kiki || '○○○○') + 'の機材設置状況及びアンカーの施工要領、引抜き強度等は、次のとおり行う。', 11, false, 22);
    line('・チャンネルベースの位置を決め、アンカーボルトの打設箇所の墨出しをする。', 11, false, 22);
    line('・据付け用アンカー（接着系アンカーの場合）の施工要領によりアンカーを打設する。', 11, false, 22);
    line('・チャンネルベースを据付け、ライナー等で水平に調整しナットで固定する。', 11, false, 22);
    line('・チャンネルベースと' + name + 'は、ボルト（○○φ）で1面当たり4箇所を固定する。', 11, false, 22);
    line('（チャンネルベースとアンカーの位置、チャンネルベースの固定の図は、この下に作成してください。）', 9, false, 18);
    r += 18;
    line('⑨ 養生方法', 12, true, 24);
    line('・' + (kiki || '機器') + 'は、搬入後ビニルシートで保護する。', 11, false, 22);
    line('・屋内型の機器は、湿気に留意し、養生を行う。', 11, false, 22);
    return ws;
  }

  /* 太陽光発電装置の施工要領（作成例 (10) の 5.1 ③④、5.2〜5.4 の本文と地域係数一覧。図は利用者が Excel で作成する） */
  var PV_ALPHA = [
    [3.6, '札幌（北海道）', '43'], [2.9, '青森、秋田、岩手', '39～40'], [2.6, '山形、宮城、福島、新潟', '37～38'],
    [2.4, '群馬、栃木、茨城、長野、石川、富山', '36'],
    [2.3, '東京、埼玉、神奈川、愛知、京都、鳥取、島根、千葉、山梨、岐阜、滋賀、兵庫、福井', '35'],
    [2.2, '静岡、大阪、広島、香川、徳島、岡山、山口、三重、奈良、和歌山', '34'],
    [2.1, '福岡、佐賀、大分、愛媛、高知', '33'], [2.0, '熊本、長崎', '32'], [1.9, '宮崎、鹿児島', '31']
  ];

  function addSuetsukePv(wb, info) {
    var ws = portraitSheet(wb, '据付け施工要領');
    ws.columns = [{ width: 10 }, { width: 66 }, { width: 14 }];
    var H = { size: 10, v: 'middle', h: 'center', fill: true };
    var C = { size: 10, v: 'middle', h: 'center' };
    var L = { size: 10, v: 'middle' };
    var INPUT = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
    var r = 1;
    function line(t, size, bold, h) { textLine(ws, r, 3, t, size, bold, h, 'left'); r++; }
    function head(t) { line(t, 12, true, 24); }
    function body(t) { line(t, 11, false, 22); }
    function note(t) { line(t, 9, false, 18); }
    textLine(ws, r, 3, '据付け施工要領', 16, true, 30); r++;
    line('工事名　' + (info.kojiName || ''), 10, false, 20);
    head('③ アレイの取付け方法');
    body('・設置方法、離隔距離などは、設計図書及び次に示す東京消防庁「太陽光発電設備に係る防火安全対策の指導基準」による。');
    body('　(1) 消防隊員が活用する施設周囲への設置抑制');
    body('　　屋外階段、非常用の進入口、代替開口部及びその周囲おおむね50cmの範囲には、PVモジュール、直流配線等を設置しないこと。');
    body('　(2) PVモジュールの屋根への設置方法');
    body('　　大規模に設置する場合、消防活動用通路を全てのPVモジュールとの距離が、24m以内となるよう設置すること。');
    body('　(3) 消防法令上の規制場所へのPVモジュールの設置');
    body('　　一定の条件を満足するPVモジュールを、屋上設備の周囲で消防法令上、建築設備等を設置できない規制場所に設置する場合は、所轄消防署と協議する。');
    body('・アレイ間の離隔距離（L）は○mとする。');
    note('　離隔距離の目安は下式にて算出　L＝H×α (m)（α＝地域係数）　H＝H2－H1 (m)　（色付きの欄に入力すると計算されます）');
    var c0 = r;
    [['前列アレイの上端の高さ H2 (m)', true], ['後列アレイの下端の高さ H1 (m)', true], ['地域係数 α（下の表から）', true],
      ['高さの差 H＝H2－H1 (m)', false], ['離隔距離の目安 L＝H×α (m)', false]].forEach(function (c, i) {
      box(ws, r, 1, r, 2, c[0], L);
      box(ws, r, 3, r, 3, '', C);
      if (c[1]) { ws.getCell(r, 3).fill = INPUT; }
      ws.getRow(r).height = 20;
      r++;
    });
    var fH = 'IF(OR(C' + c0 + '="",C' + (c0 + 1) + '=""),"",C' + c0 + '-C' + (c0 + 1) + ')';
    var fL = 'IF(OR(C' + (c0 + 3) + '="",C' + (c0 + 2) + '=""),"",ROUND(C' + (c0 + 3) + '*C' + (c0 + 2) + ',2))';
    ws.getCell(c0 + 3, 3).value = { formula: fH, result: '' };
    ws.getCell(c0 + 4, 3).value = { formula: fL, result: '' };
    ws.getRow(r).height = 8; r++;
    line('地域係数一覧（真南：冬至：午前9時及び午後3時）', 10, true, 20);
    box(ws, r, 1, r, 1, 'α', H); box(ws, r, 2, r, 2, '地域', H); box(ws, r, 3, r, 3, '北緯（°）', H);
    ws.getRow(r).height = 20; r++;
    PV_ALPHA.forEach(function (a) {
      box(ws, r, 1, r, 1, a[0], C);
      ws.getCell(r, 1).numFmt = '0.0';
      box(ws, r, 2, r, 2, a[1], L);
      box(ws, r, 3, r, 3, a[2], C);
      ws.getRow(r).height = Math.max(20, lineCount(a[1], 66, 10) * 13.5 + 6);
      r++;
    });
    ws.getRow(r).height = 8; r++;
    body('・アンカーボルトと架台の取付けは、次のとおりとする。');
    note('（アレイの離隔距離の図、アンカーボルトと架台の取付けの図は、別のシートまたはこの下に作成してください。）');
    head('④ 養生');
    body('・アレイは搬入後、架台に取付けるまではビニルシートで保護する。');
    body('・太陽電池モジュールは太陽光にあたると発電を開始するので、取付け工事・配線工事中は、ビニール等で充電部の養生を行い感電を防止する。');
    head('5.2 アレイを設置する構造物の条件');
    body('・太陽電池アレイ用支持物等は、JIS C 8955：2017「太陽電池アレイ用支持物の設計用荷重算出方法」による想定荷重（固定荷重、風圧荷重、積雪荷重、地震荷重）の計算を行い、想定荷重に耐える強度の支持構造物及び支持ボルト等を算出し、施工図又は機器製作図に記載する。');
    head('5.3 パワーコンディショナの設置方法');
    body('・パワーコンディショナの重量を支える強度のある壁面に堅固に据付ける。');
    body('・PV直流用SPDは、接続箱とパワーコンディショナ内に設ける。');
    note('（平面図・正面図は別に作成してください。図中寸法は、製造者仕様書等を参考にする。）');
    head('5.4 表示');
    body('・消防法により、接続箱、パワーコンディショナ及びアレイからパワーコンディショナまでの直流配線、配管に「太陽光直流配線注意」の表示を行う。');
    body('・発電用太陽電池設備に関する技術基準を定める省令及び同解釈により、接触防止・立入防止の注意標識を設ける。（設置場所は監督職員と協議のうえ決定する。）');
    return ws;
  }

  /* 据付け用アンカー（接着系アンカーの場合）の施工要領（作成例 (8) p.439 の全文） */
  function addAnchor(wb, info) {
    var ws = portraitSheet(wb, 'アンカーの施工要領');
    ws.columns = [{ width: 10 }, { width: 44 }, { width: 10 }, { width: 28 }];
    var H = { size: 10, v: 'middle', h: 'center', fill: true };
    var V = { size: 10, v: 'middle' };
    textLine(ws, 1, 4, '据付け用アンカー（接着系アンカーの場合）の施工要領', 14, true, 30);
    box(ws, 2, 1, 2, 1, '工事名', H);
    box(ws, 2, 2, 2, 4, info.kojiName || '', V);
    box(ws, 3, 1, 3, 1, '作成者', H);
    box(ws, 3, 2, 3, 2, info.sakuseisha || '', V);
    box(ws, 3, 3, 3, 3, '日付', H);
    box(ws, 3, 4, 3, 4, wareki(info.sakuseiBi), V);
    ws.getRow(2).height = 22;
    ws.getRow(3).height = 22;
    ws.getRow(4).height = 8;
    var T = [
      ['1. アンカー位置の墨出し作業手順', [
        'キュービクル用コンクリート基礎工事前にチャンネルベース固定用アンカーの墨出しを行う。',
        '型枠、鉄筋工事が完了後、アンカー位置が鉄筋と干渉していないかを確認する。',
        '干渉している箇所については、鉄筋を緩やかに曲げ、アンカーの下穴を開けたときに、鉄筋とぶつからないように処置をしておく。',
        '曲げた鉄筋は必要に応じて補強筋を施しておく。',
        'コンクリート基礎打設後、型枠をばらし、基礎上の防じん塗装が完了後、アンカーの位置の墨出しを行う。'
      ]],
      ['2. 接着系アンカーの施工手順', [
        '墨出し位置に施工面に対し垂直方向に、仕様にあった適正なドリルで穴あけを行う。',
        '孔内の切粉はブロアーで取り除いたあと、専用ブラシで内壁に付着している切粉を落とし、再度ブロアーで切粉を取り除く。',
        'アンカーボルトにマーキングを行う。',
        'カプセルを孔底まで挿入する。',
        '1.0～1.5kgの重量のハンマードリルでボルトを孔底まで打ち込む。',
        '硬化するまで養生する。'
      ]],
      ['3. 注意事項', [
        'ドリルのビット径が適切であるか、必ず確認を行う。',
        '孔内の清掃を行う場合は金属ワイヤブラシ等の強度のあるものを使用し、切粉の落としを丁寧に行う。',
        '孔内に水がある場合は必ず取り除くこと。',
        'ボルトを孔底まで確実に埋め込むためにマーキングを行い、打ち込み後確認を行う。',
        'ボルトに油分が付着している場合はきれいに拭き取る。',
        'ボルトの頭部にはナットをかける等の保護を行う。',
        'カプセルの使用期限と樹脂の流動性があることを確認し、使用期限が切れているものや流動性がないものは使用しない。',
        '硬化剤が欠落したカプセルは使用しない。',
        'カプセルが孔からはみ出る場合は、はみ出た部分をハンマー等で叩き割り、その部分も孔内に入れる。その際はガラス片を含むので、素手ではさわらない。',
        'カプセルは必ず孔底まで挿入する。',
        'ボルトを打ち込む際は、必ずストッパーを使用する。',
        '適正な重量のハンマードリルを使用する。',
        '衝撃的に打ち込むと液飛散の原因となるので5～6回以上で均一に孔底まで打ち込む。',
        '樹脂があふれることを確認する。',
        '硬化時間内はボルトを動かさない。',
        '施工に用いたナット類は硬化時間を過ぎてから取外す。'
      ]]
    ];
    var r = 5;
    T.forEach(function (sec) {
      textLine(ws, r, 4, sec[0], 11, true, 22, 'left');
      r++;
      sec[1].forEach(function (t, i) {
        textLine(ws, r, 4, '(' + (i + 1) + ') ' + t, 10, false, 18, 'left');
        r++;
      });
      ws.getRow(r).height = 8;
      r++;
    });
    return ws;
  }

  // ほかのファイルから使う部品
  X.hannyuKiki = hannyuKiki;
  X.addHannyu = addHannyu;
  X.addHannyuTaisei = addHannyuTaisei;
  X.addHannyuZu = addHannyuZu;
  X.WIRE_T1 = WIRE_T1;
  X.WIRE_T2 = WIRE_T2;
  X.addWire = addWire;
  X.addSuetsuke = addSuetsuke;
  X.PV_ALPHA = PV_ALPHA;
  X.addSuetsukePv = addSuetsukePv;
  X.addAnchor = addAnchor;
})(window.SKUKIKKSExport._);
