/* 施工計画書 Excel(xlsx) 出力：表紙・工事概要・組織表・電気保安体制・品質管理体制・総合施工計画書の添付シート
 * 分割したファイルは index.html の順（core → … → main）に読み込む。部品は window.SekokanExport._ で共有する
 */
(function (X) {
  'use strict';

  var FONT = X.FONT, KOSHU = X.KOSHU, box = X.box, hline = X.hline, landscapeSheet = X.landscapeSheet,
    lineCount = X.lineCount, orMaru = X.orMaru, portraitSheet = X.portraitSheet, textLine = X.textLine,
    totalWidth = X.totalWidth, vline = X.vline, wareki = X.wareki;

  /* ---------- 表紙（書籍 第Ⅲ編 2.2 総合施工計画書作成例の表紙にならう） ---------- */

  // koshu が true のとき（工種別）は「管理技術者」の ※ と注記を付けない（作成例 p.345。K-39）
  function addCover(wb, info, koshu) {
    var ws = portraitSheet(wb, '表紙');
    // B〜E の4列を同じ幅にして、押印欄（4人分）に使う
    ws.columns = [{ width: 3 }, { width: 19.5 }, { width: 19.5 }, { width: 19.5 }, { width: 19.5 }, { width: 3 }];
    var LAST = 6;
    var HEADOPT = { size: 11, v: 'middle', h: 'center', fill: true };
    var DATEOPT = { size: 10, v: 'middle', h: 'center' };
    var DATE = '年　　月　　日';
    var STAMP = 64;        // 押印スペースの高さ(pt)
    var c;

    ws.getRow(1).height = 50;
    textLine(ws, 2, LAST, '工事名　　' + (info.kojiName || ''), 18, false, 36);
    ws.getRow(3).height = 50;
    // 書類名が1行に収まらないときは文字を小さくする（工種別は書類名が長い）
    var title = info.docTitle || '施工計画書';
    var tsize = 34;
    if (lineCount(title, totalWidth(ws, LAST), tsize) > 1) { tsize = 26; }
    if (lineCount(title, totalWidth(ws, LAST), tsize) > 1) { tsize = 20; }
    textLine(ws, 4, LAST, title, tsize, true, 64);
    ws.getRow(5).height = 70;

    textLine(ws, 6, LAST, '　工事監理　' + (info.kojiKanri || ''), 12, false, 22, 'left');
    textLine(ws, 7, LAST, '　監　　理　' + (info.kanri || ''), 12, false, 22, 'left');
    ws.getRow(8).height = 8;

    var r = 9;
    box(ws, r, 2, r, 3, '受理・承諾', HEADOPT);
    box(ws, r, 4, r, 5, '確認', HEADOPT);
    ws.getRow(r).height = 22;
    r++;
    ['主任監督員', '監督員', koshu ? '管理技術者' : '管理技術者※', '担当者'].forEach(function (t, i) {
      box(ws, r, 2 + i, r, 2 + i, t, HEADOPT);
    });
    ws.getRow(r).height = 22;
    r++;
    for (c = 2; c <= 5; c++) { box(ws, r, c, r, c, DATE, DATEOPT); }
    ws.getRow(r).height = 22;
    r++;
    // 受理・承諾側は押印欄の下に注記の行が入る。確認側は注記の行まで押印欄にする
    box(ws, r, 2, r, 2, '', {});
    box(ws, r, 3, r, 3, '', {});
    box(ws, r, 4, r + 1, 4, '', {});
    box(ws, r, 5, r + 1, 5, '', {});
    ws.getRow(r).height = STAMP - 18;
    r++;
    box(ws, r, 2, r, 3, '承諾は品質計画に係る部分', { size: 9, v: 'middle', h: 'center' });
    ws.getRow(r).height = 18;
    r++;
    if (!koshu) {
      textLine(ws, r, LAST, '　※「管理技術者」は、工事監理業務の受託者の代表となる工事監理者', 9, false, 16, 'left');
      r++;
    }

    ws.getRow(r).height = 36;
    r++;
    textLine(ws, r, LAST, '　受注者　' + (info.juchu || ''), 12, false, 22, 'left');
    r++;
    ws.getRow(r).height = 8;
    r++;
    ['現場代理人', '監理(主任)技術者', '工事担当者', '工事担当者'].forEach(function (t, i) {
      box(ws, r, 2 + i, r, 2 + i, t, HEADOPT);
    });
    ws.getRow(r).height = 22;
    r++;
    for (c = 2; c <= 5; c++) { box(ws, r, c, r, c, DATE, DATEOPT); }
    ws.getRow(r).height = 22;
    r++;
    for (c = 2; c <= 5; c++) { box(ws, r, c, r, c, '', {}); }
    ws.getRow(r).height = STAMP;
    r++;

    ws.getRow(r).height = 30;
    r++;
    var tail = [];
    if (info.sakuseiBi) { tail.push('作成年月日　' + wareki(info.sakuseiBi)); }
    if (info.sakuseisha) { tail.push('作成者　' + info.sakuseisha); }
    if (tail.length) { textLine(ws, r, LAST, tail.join('　　　') + '　', 11, false, 22, 'right'); }
    return ws;
  }

  /* ---------- 工事概要（作成例「2. 工事概要」①〜⑯にならう） ---------- */

  // 選択した工事種目(章)から、電気設備概要の表に並べる設備種目を決める
  var SETSUBI = {
    3: ['電灯設備', '動力設備'], 4: ['受変電設備'], 5: ['電力貯蔵設備'], 6: ['発電設備'],
    7: ['通信・情報設備'], 8: ['中央監視制御設備'], 9: ['医療関係設備']
  };

  function addGaiyo(wb, info, chapters, withKanri) {
    var ws = portraitSheet(wb, '工事概要');
    ws.columns = [{ width: 4 }, { width: 18 }, { width: 18 }, { width: 18 }, { width: 28 }];
    var LAST = 5;
    var HEADOPT = { size: 10, v: 'middle', h: 'center', fill: true };
    var CELLOPT = { size: 10, v: 'middle' };

    textLine(ws, 1, LAST, '工事概要', 16, true, 30);
    textLine(ws, 2, LAST, '当該工事の概要を下記①～⑯に記す。', 11, false, 22, 'left');

    var koki = (info.kokiFrom || info.kokiTo)
      ? ('自 ' + wareki(info.kokiFrom) + '　～　至 ' + wareki(info.kokiTo)) : '';
    var rows = [
      ['①', '工事名称', info.kojiName],
      ['②', '工事場所', info.basho],
      ['③', '発注者', info.hacchu],
      ['④', '工事監理', info.kojiKanri],
      ['⑤', '監理', info.kanri],
      ['⑥', '受注者', info.juchu],
      ['⑦', '設計者', info.sekkei],
      ['⑧', '工期', koki],
      ['⑨', '建物用途', info.yoto],
      ['⑩', '地域地区等', info.chiiki],
      ['⑪', '構造', info.kozo],
      ['⑫', '規模', info.kibo],
      ['⑬', '消防法上の区分', info.shobo],
      ['⑭', '施設の分類', info.bunrui]
    ];
    var r = 3;
    rows.forEach(function (row) {
      box(ws, r, 1, r, 1, row[0], { size: 11, v: 'middle', h: 'center' });
      box(ws, r, 2, r, 2, row[1], { size: 11, v: 'middle', fill: true });
      box(ws, r, 3, r, 5, row[2] || '', { size: 11, v: 'middle' });
      ws.getRow(r).height = Math.max(24, lineCount(row[2] || '', 62, 11) * 16 + 8);
      r++;
    });

    ws.getRow(r).height = 10;
    r++;
    textLine(ws, r, LAST, '⑮　電気設備概要', 11, false, 22, 'left');
    r++;
    box(ws, r, 1, r, 2, '設備種目', HEADOPT);
    ['方式等', '機器名等', '形式、容量等'].forEach(function (t, i) {
      box(ws, r, 3 + i, r, 3 + i, t, HEADOPT);
    });
    ws.getRow(r).height = 20;
    r++;
    var setsubi = [];
    chapters.forEach(function (ch) {
      (SETSUBI[ch.id] || []).forEach(function (n) { setsubi.push(n); });
    });
    if (!setsubi.length) { setsubi = ['', '', '']; }
    setsubi.forEach(function (n) {
      box(ws, r, 1, r, 2, n, CELLOPT);
      box(ws, r, 3, r, 3, '', CELLOPT);
      box(ws, r, 4, r, 4, '', CELLOPT);
      box(ws, r, 5, r, 5, '', CELLOPT);
      ws.getRow(r).height = 26;
      r++;
    });

    ws.getRow(r).height = 10;
    r++;
    textLine(ws, r, LAST, '⑯　機械設備概要', 11, false, 22, 'left');
    r++;
    box(ws, r, 1, r, 2, '主設備', HEADOPT);
    box(ws, r, 3, r, 3, '方式等', HEADOPT);
    box(ws, r, 4, r, 5, '主要機器名等', HEADOPT);
    ws.getRow(r).height = 20;
    r++;
    var k;
    for (k = 0; k < 2; k++) {
      box(ws, r, 1, r, 2, '', CELLOPT);
      box(ws, r, 3, r, 3, '', CELLOPT);
      box(ws, r, 4, r, 5, '', CELLOPT);
      ws.getRow(r).height = 26;
      r++;
    }

    if (!chapters.length || !withKanri) { return ws; }
    ws.getRow(r).height = 14;
    r++;
    textLine(ws, r, LAST, '管理表（品質計画）の対象工事種目', 11, false, 22, 'left');
    r++;
    var names = chapters.map(function (ch) { return ch.title.replace(/　/g, ' '); });
    box(ws, r, 1, r, 5, names.join('　／　'), { size: 10, v: 'middle' });
    ws.getRow(r).height = Math.max(24, lineCount(names.join('　／　'), 84, 10) * 15 + 8);
    r++;
    textLine(ws, r, LAST,
      '管理表は、国土交通省「電気設備工事 施工管理 管理表（令和8年版）」から対象の工事種目を抜粋したものです。',
      9, false, 20, 'left');
    return ws;
  }

  /* ---------- 組織表・電気保安体制（作成例「3. 受注者の組織」にならう） ---------- */


  function addSoshiki(wb, info) {
    var ws = portraitSheet(wb, '組織表');
    var cols = [];
    var i;
    for (i = 0; i < 14; i++) { cols.push({ width: 6.2 }); }
    ws.columns = cols;
    var LAST = 14;
    var B = { size: 11, v: 'middle', h: 'center' };
    var H = 20;
    var r;

    textLine(ws, 1, LAST, '受注者の組織', 16, true, 30);
    textLine(ws, 2, LAST, '当該工事の組織表を次に示す。', 11, false, 22, 'left');
    textLine(ws, 3, LAST, '社内組織', 11, true, 22, 'left');

    box(ws, 4, 5, 5, 10, orMaru(info.juchu, '○○○○電気工事株式会社') + '\n代表取締役社長　○○ ○○', B);
    vline(ws, 6, 6, 7);
    hline(ws, 7, 3, 12);
    vline(ws, 7, 7, 2); vline(ws, 7, 7, 7); vline(ws, 7, 7, 12);
    box(ws, 8, 1, 10, 4, '○○部\n部長　○○ ○○\n課長　○○ ○○', B);
    box(ws, 8, 6, 10, 9, '○○部\n部長　○○ ○○\n課長　○○ ○○', B);
    box(ws, 8, 11, 10, 14, '○○部\n部長　○○ ○○', B);
    vline(ws, 11, 11, 7);

    // 「現場の組織」の見出し行。縦線を通すため結合は左4列だけにする
    ws.mergeCells(12, 1, 12, 4);
    ws.getCell(12, 1).value = '現場の組織';
    ws.getCell(12, 1).font = { name: FONT, size: 11, bold: true };
    ws.getCell(12, 1).alignment = { vertical: 'middle' };
    vline(ws, 12, 12, 7);

    box(ws, 13, 6, 14, 9, '現場代理人\n' + orMaru(info.dairinin), B);
    vline(ws, 15, 15, 7);
    box(ws, 16, 1, 17, 5, '監理(主任)技術者（注）\n' + orMaru(info.gijutsu), B);
    box(ws, 16, 10, 17, 14, '電気保安技術者\n○○ ○○', B);
    hline(ws, 17, 6, 9);
    vline(ws, 16, 17, 7);
    vline(ws, 18, 18, 7);
    hline(ws, 19, 4, 11);
    vline(ws, 19, 19, 3); vline(ws, 19, 19, 11);
    box(ws, 20, 2, 21, 5, '○○担当者\n○○ ○○', B);
    box(ws, 20, 10, 21, 13, '○○担当者\n○○ ○○', B);
    vline(ws, 22, 22, 11);
    box(ws, 23, 10, 24, 13, '職長　○○ ○○', B);
    box(ws, 23, 1, 24, 8,
      '（注）建設業法第26条第3項第2号の規定により監理技術者補佐を専任で置く場合は組織表に記載する。',
      { size: 9, v: 'middle' });

    for (r = 4; r <= 24; r++) { ws.getRow(r).height = H; }
    ws.getRow(12).height = 24;
    ws.getRow(25).height = 10;
    textLine(ws, 26, LAST, '※下請業者との関係は、「施工体制台帳」で管理する。', 10, false, 20, 'left');
    textLine(ws, 27, LAST, '（○○ の箇所は、部署名・氏名に書き換えてください。枠はセルの罫線です。）', 9, false, 18, 'left');
    return ws;
  }

  function addHoan(wb) {
    var ws = portraitSheet(wb, '電気保安体制');
    var cols = [];
    var i;
    for (i = 0; i < 14; i++) { cols.push({ width: 6.2 }); }
    ws.columns = cols;
    var LAST = 14;
    var B = { size: 11, v: 'middle', h: 'center' };
    var r;

    textLine(ws, 1, LAST, '電気保安体制', 16, true, 30);
    ws.getRow(2).height = 14;
    box(ws, 3, 5, 4, 10, '電気主任技術者\n（発注者）', B);
    vline(ws, 5, 6, 7);
    box(ws, 7, 5, 8, 10, '工事担当技術者\n（監督職員）', B);
    vline(ws, 9, 10, 7);
    box(ws, 11, 5, 12, 10, '電気保安技術者\n○○ ○○', B);
    for (r = 3; r <= 12; r++) { ws.getRow(r).height = 20; }
    ws.getRow(13).height = 14;
    textLine(ws, 14, LAST,
      '電気保安技術者は、発注者側で定めた電気主任技術者の業務を補佐する監督職員の指示に従い、' +
      '電気工作物の保安業務を行う。（「標仕」第1編 1.3.2）', 10, false, 34, 'left');
    return ws;
  }

  /* 品質管理体制（作成例「7.1 品質管理体制」の図にならう） */
  function addHinshitsu(wb, info) {
    var ws = portraitSheet(wb, '品質管理体制');
    var cols = [];
    var i;
    for (i = 0; i < 14; i++) { cols.push({ width: 6.2 }); }
    ws.columns = cols;
    var LAST = 14;
    var B = { size: 11, v: 'middle', h: 'center' };
    var r;

    textLine(ws, 1, LAST, '品質管理体制', 16, true, 30);
    textLine(ws, 2, LAST, '当該工事の品質管理体制を次に示す。', 11, false, 22, 'left');

    box(ws, 3, 8, 4, 13, '現場代理人\n' + orMaru(info.dairinin), B);
    hline(ws, 4, 4, 7);                 // 現場代理人の左から監理技術者の上へ
    vline(ws, 4, 4, 3);
    box(ws, 5, 1, 6, 6, '監理(主任)技術者\n' + orMaru(info.gijutsu), B);
    vline(ws, 7, 7, 3);
    hline(ws, 8, 4, 10);
    vline(ws, 8, 8, 3); vline(ws, 8, 8, 10);
    box(ws, 9, 1, 10, 6, '○○担当者\n○○ ○○', B);
    box(ws, 9, 8, 10, 13, '○○担当者\n○○ ○○', B);
    vline(ws, 11, 11, 3); vline(ws, 11, 11, 10);
    box(ws, 12, 1, 14, 6, '主任技術者\n○○○電気工事\n○○ ○○', B);
    box(ws, 12, 8, 14, 13, '主任技術者\n○○○電気工事\n○○ ○○', B);
    vline(ws, 15, 15, 3); vline(ws, 15, 15, 10);
    box(ws, 16, 1, 17, 6, '職長　○○ ○○', B);
    box(ws, 16, 8, 17, 13, '職長　○○ ○○', B);

    for (r = 3; r <= 17; r++) { ws.getRow(r).height = 20; }
    ws.getRow(18).height = 10;
    textLine(ws, 19, LAST, '（○○ の箇所は、会社名・氏名に書き換えてください。枠はセルの罫線です。）', 9, false, 18, 'left');
    return ws;
  }

  /* ---------- 総合施工計画書の添付シート ---------- */

  function addKoshu(wb) {
    var ws = portraitSheet(wb, '工種別施工計画書一覧');
    ws.columns = [{ width: 6 }, { width: 62 }, { width: 18 }];
    textLine(ws, 1, 3, '工種別施工計画書一覧', 16, true, 30);
    textLine(ws, 2, 3, '工種別施工計画書は、次表に示す工種について作成する。（作成しない行は削除してください。）',
      10, false, 22, 'left');
    var H = { size: 11, v: 'middle', h: 'center', fill: true };
    box(ws, 3, 1, 3, 1, 'No.', H);
    box(ws, 3, 2, 3, 2, '工種別施工計画書', H);
    box(ws, 3, 3, 3, 3, '提出予定', H);
    ws.getRow(3).height = 22;
    KOSHU.forEach(function (name, i) {
      var r = 4 + i;
      box(ws, r, 1, r, 1, i + 1, { size: 11, v: 'middle', h: 'center' });
      box(ws, r, 2, r, 2, name, { size: 11, v: 'middle' });
      box(ws, r, 3, r, 3, '', { size: 11, v: 'middle' });
      ws.getRow(r).height = 24;
    });
    return ws;
  }

  /* 安全衛生管理組織（作成例 8.3 の図） */
  function addAnzenSoshiki(wb, info) {
    var ws = portraitSheet(wb, '安全衛生管理組織');
    var cols = [];
    var i;
    for (i = 0; i < 19; i++) { cols.push({ width: 5.6 }); }
    ws.columns = cols;
    var LAST = 19;
    var B = { size: 10, v: 'middle', h: 'center' };
    var S = { size: 9, v: 'middle', h: 'center' };
    var r;

    textLine(ws, 1, LAST, '安全衛生管理組織', 16, true, 30);
    textLine(ws, 2, LAST, '安全衛生管理組織を次表に示す。', 11, false, 22, 'left');

    box(ws, 3, 6, 4, 9, '統括安全衛生責任者', B);
    hline(ws, 4, 10, 12);
    box(ws, 3, 13, 4, 16, '災害防止協議会', B);
    ws.mergeCells(5, 1, 5, 5);
    ws.getCell(5, 1).value = '建築工事業者（元請）';
    ws.getCell(5, 1).font = { name: FONT, size: 10 };
    ws.getCell(5, 1).alignment = { vertical: 'middle' };
    vline(ws, 5, 5, 7);
    box(ws, 6, 6, 7, 9, '元方安全衛生管理者', B);
    vline(ws, 8, 8, 7);
    hline(ws, 9, 3, 14);
    vline(ws, 9, 9, 2); vline(ws, 9, 9, 7); vline(ws, 9, 9, 14);
    box(ws, 10, 12, 11, 17, orMaru(info.juchu, '○○○○電気工事株式会社') + '（元請）\n安全衛生責任者', S);
    vline(ws, 10, 13, 2); vline(ws, 10, 13, 7);
    vline(ws, 12, 12, 14);
    hline(ws, 13, 13, 17);
    vline(ws, 13, 13, 12); vline(ws, 13, 13, 17);
    box(ws, 14, 1, 15, 4, '建築業者（とび）\n一次下請　安全衛生責任者', S);
    box(ws, 14, 6, 15, 9, '建築業者（大工）\n一次下請　安全衛生責任者', S);
    box(ws, 14, 11, 15, 14, '電気工事業者（電力）\n一次下請　安全衛生責任者', S);
    box(ws, 14, 16, 15, 19, '電気工事業者（通信）\n一次下請　安全衛生責任者', S);
    vline(ws, 16, 16, 2); vline(ws, 16, 16, 7); vline(ws, 16, 16, 12); vline(ws, 16, 16, 17);
    box(ws, 17, 1, 18, 4, '建築業者（とび）\n二次下請　安全衛生責任者', S);
    box(ws, 17, 6, 18, 9, '建築業者（大工）\n二次下請　安全衛生責任者', S);
    box(ws, 17, 11, 18, 14, '電気工事業者（電力）\n二次下請　安全衛生責任者', S);
    box(ws, 17, 16, 18, 19, '電気工事業者（通信）\n二次下請　安全衛生責任者', S);

    for (r = 3; r <= 18; r++) { ws.getRow(r).height = 20; }
    ws.getRow(19).height = 10;
    textLine(ws, 20, LAST, '（業者名・職種は当該工事に合わせて書き換えてください。枠はセルの罫線です。）', 9, false, 18, 'left');
    return ws;
  }

  /* 安全衛生管理活動の計画表（作成例 8.4）。工期が入力されていれば、その月を並べる */
  function monthList(from, to) {
    var a = /^(\d{4})-(\d{2})/.exec(from || '');
    var b = /^(\d{4})-(\d{2})/.exec(to || '');
    var list = [];
    var i;
    if (a && b) {
      var y = Number(a[1]);
      var m = Number(a[2]);
      var ey = Number(b[1]);
      var em = Number(b[2]);
      while ((y < ey || (y === ey && m <= em)) && list.length < 60) {
        list.push({ y: y, m: m });
        m++;
        if (m > 12) { m = 1; y++; }
      }
    }
    if (!list.length) {
      for (i = 0; i < 18; i++) { list.push({ y: 0, m: 0 }); }
    }
    return list;
  }

  function addAnzenKeikaku(wb, info) {
    var ws = landscapeSheet(wb, '安全衛生活動計画表');
    var months = monthList(info.kokiFrom, info.kokiTo);
    var n = months.length;
    var cols = [{ width: 22 }];
    var i;
    for (i = 0; i < n; i++) { cols.push({ width: Math.max(4.2, Math.min(7, 100 / n)) }); }
    ws.columns = cols;
    var LAST = n + 1;
    var H = { size: 10, v: 'middle', h: 'center', fill: true };
    var C = { size: 10, v: 'middle', h: 'center' };

    textLine(ws, 1, LAST, '安全衛生管理活動の計画表', 16, true, 30);
    textLine(ws, 2, LAST, '安全衛生管理活動の計画を次表に示す。', 11, false, 22, 'left');

    box(ws, 3, 1, 4, 1, '項目　＼　工程', H);
    // 年の見出し（同じ年の月をまとめる）
    var start = 0;
    while (start < n) {
      var end = start;
      while (end + 1 < n && months[end + 1].y === months[start].y) { end++; }
      var ry = months[start].y - 2018;
      var label = '令和　年';
      if (months[start].y) {
        label = ry >= 1 ? ('令和' + (ry === 1 ? '元' : ry) + '年') : (months[start].y + '年');
        // 1か月分しか幅がないときは短い表記にする
        if (end === start) { label = ry >= 1 ? ('R' + ry) : String(months[start].y).substring(2); }
      }
      box(ws, 3, 2 + start, 3, 2 + end, label, H);
      start = end + 1;
    }
    months.forEach(function (mm, k) {
      box(ws, 4, 2 + k, 4, 2 + k, mm.m ? (mm.m + '月') : '月', H);
    });
    ws.getRow(3).height = 20;
    ws.getRow(4).height = 20;

    var rows = [
      ['特別安全朝礼\n（毎月1日）', function () { return true; }],
      ['安全パトロール\n（随時）', function () { return false; }],
      ['災害防止協議会\n（第2金曜日）', function () { return true; }],
      ['安全大会\n（7/1～、12/1～）', function (mm) { return mm.m === 7 || mm.m === 12; }]
    ];
    rows.forEach(function (row, ri) {
      var r = 5 + ri;
      box(ws, r, 1, r, 1, row[0], C);
      months.forEach(function (mm, k) {
        box(ws, r, 2 + k, r, 2 + k, row[1](mm) ? '○' : '', C);
      });
      ws.getRow(r).height = 32;
    });
    textLine(ws, 10, LAST, '（○ は実施予定月。安全パトロール等は当該工事の計画に合わせて記入してください。）', 9, false, 18, 'left');
    return ws;
  }

  /* 事故時等の連絡網（作成例 8.7）。連絡先の表を並べ、監督職員への連絡順は縦線でつなぐ */
  function addRenraku(wb) {
    var ws = portraitSheet(wb, '事故時等の連絡網');
    ws.columns = [{ width: 10 }, { width: 17 }, { width: 3 }, { width: 10 }, { width: 17 }, { width: 3 },
      { width: 10 }, { width: 17 }];
    var H = { size: 10, v: 'middle', fill: true };
    var C = { size: 10, v: 'middle' };

    // 見出し＋項目行の表を置き、次の行番号を返す
    function block(r, c, title, fields) {
      box(ws, r, c, r, c + 1, title, H);
      fields.forEach(function (f, i) {
        box(ws, r + 1 + i, c, r + 1 + i, c, f, C);
        box(ws, r + 1 + i, c + 1, r + 1 + i, c + 1, '', C);
      });
      return r + 1 + fields.length;
    }

    textLine(ws, 1, 8, '事故時等の連絡網', 16, true, 30);
    textLine(ws, 2, 8, '事故時等の連絡先・体制図を次に示す。', 11, false, 22, 'left');

    // 左: 関係官公署、所属部局、各課・室
    box(ws, 3, 1, 3, 2, '関係官公署', { size: 10, v: 'middle', h: 'center', fill: true });
    ['消防署', '労基署', '病院', '警察署', '電力会社', '水道局', '下水道局', 'ガス会社', '電話会社', ''].forEach(function (nm, i) {
      box(ws, 4 + i, 1, 4 + i, 1, nm, C);
      box(ws, 4 + i, 2, 4 + i, 2, 'TEL', C);
    });
    block(16, 1, '4. 所属部局', ['TEL', 'FAX']);
    vline(ws, 19, 19, 1);
    block(20, 1, '5. 各課、室', ['TEL', 'FAX', '', '']);

    // 中央: 現場代理人 → 1.主任監督員 → 2.総括監督員 → 3.監督員
    block(3, 4, '現場代理人', ['担当', 'TEL', 'FAX', '自宅']);
    vline(ws, 8, 8, 4);
    block(9, 4, '1. 主任監督員', ['担当', 'TEL', 'FAX', '自宅']);
    vline(ws, 14, 14, 4);
    block(15, 4, '2. 総括監督員', ['担当', '自宅']);
    vline(ws, 18, 18, 4);
    block(19, 4, '3. 監督員', ['担当', '自宅']);

    // 右: 会社、協力会社、関係業者、監理業務受注者、相手庁
    block(3, 7, '会社', ['担当', 'TEL', '夜間']);
    block(8, 7, '協力会社', ['担当', 'TEL', '夜間']);
    block(13, 7, '関係業者', ['担当', 'TEL', '夜間']);
    block(18, 7, '監理業務受注者', ['担当', 'TEL']);
    block(22, 7, '相手庁', ['担当', 'TEL']);

    var r;
    for (r = 3; r <= 25; r++) { ws.getRow(r).height = 20; }
    ws.getRow(26).height = 10;
    textLine(ws, 27, 8, '緊急時は 1.～5. の順位で速やかに連絡を行う。', 11, true, 22, 'left');
    textLine(ws, 28, 8, '（現場代理人から関係官公署・会社・協力会社等への連絡線は省略しています。）', 9, false, 18, 'left');
    return ws;
  }

  // ほかのファイルから使う部品
  X.addCover = addCover;
  X.SETSUBI = SETSUBI;
  X.addGaiyo = addGaiyo;
  X.addSoshiki = addSoshiki;
  X.addHoan = addHoan;
  X.addHinshitsu = addHinshitsu;
  X.addKoshu = addKoshu;
  X.addAnzenSoshiki = addAnzenSoshiki;
  X.monthList = monthList;
  X.addAnzenKeikaku = addAnzenKeikaku;
  X.addRenraku = addRenraku;
})(window.SekokanExport._);
