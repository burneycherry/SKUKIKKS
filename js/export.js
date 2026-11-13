/* 施工計画書 Excel(xlsx) 出力
 * 依存: vendor/exceljs.min.js (window.ExcelJS)
 * 公開: window.SekokanExport = { build, collect, parseDocs, wareki, KOSHU }
 */
(function () {
  'use strict';

  var FONT = 'ＭＳ Ｐゴシック';
  var FS = 8;            // 管理表の文字サイズ(pt)
  var LH = 10.5;          // 1行あたりの高さ(pt)
  var CHUNK = 420;       // 区分・項目の結合を区切る高さ(pt)。ページをまたいでも見出しが出るようにする
  var THIN = { style: 'thin', color: { argb: 'FF000000' } };
  var BORDER = { top: THIN, left: THIN, bottom: THIN, right: THIN };
  var HEAD_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCE6F1' } };

  // 管理表の列幅 (A〜N)
  var W = [3, 9, 7, 7, 9, 30, 22, 20, 3.4, 16, 9, 4, 4, 5];

  /* ---------- 文字列ユーティリティ ---------- */

  // 見た目の長さ (半角=1, 全角=2)
  function vlen(s) {
    var n = 0;
    for (var i = 0; i < s.length; i++) {
      var c = s.charCodeAt(i);
      n += (c < 0x7f || (c >= 0xff61 && c <= 0xff9f)) ? 1 : 2;
    }
    return n;
  }

  // 折り返し後の行数を見積もる
  function lineCount(text, width, fontSize) {
    if (!text) { return 0; }
    var cap = Math.max(2, Math.floor(width * 1.25 * (FS / (fontSize || FS)) - 0.5));
    var lines = String(text).split('\n');
    var n = 0;
    for (var i = 0; i < lines.length; i++) {
      n += Math.max(1, Math.ceil(vlen(lines[i]) / cap));
    }
    return n;
  }

  function heightFor(lines, lh) {
    return Math.max(15, lines * (lh || LH) + 4);
  }

  // 西暦日付(yyyy-mm-dd) → 令和表記
  function wareki(iso) {
    if (!iso) { return ''; }
    var m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(iso);
    if (!m) { return iso; }
    var y = Number(m[1]);
    var mo = Number(m[2]);
    var d = Number(m[3]);
    if (y > 2019 || (y === 2019 && mo >= 5)) {
      var r = y - 2018;
      return '令和' + (r === 1 ? '元' : r) + '年' + mo + '月' + d + '日';
    }
    return y + '年' + mo + '月' + d + '日';
  }

  // 作成資料欄(□で始まる箇条書き)を資料名の配列に分解
  function parseDocs(text) {
    var out = [];
    if (!text) { return out; }
    var lines = String(text).split('\n');
    for (var i = 0; i < lines.length; i++) {
      var ln = lines[i].replace(/^[\s　]+/, '').replace(/[\s　]+$/, '');
      if (!ln) { continue; }
      if (ln.charAt(0) === '□') {
        out.push(ln.substring(1));
      } else if (out.length) {
        out[out.length - 1] += ln;   // 折り返しの続き
      } else {
        out.push(ln);
      }
    }
    return out;
  }

  /* ---------- 選択内容の収集 ---------- */

  // 選択された細目だけを残した章の配列を返す
  function collect(data, sel) {
    var chapters = [];
    data.chapters.forEach(function (ch) {
      var items = [];
      ch.items.forEach(function (it) {
        var groups = it.groups.filter(function (g) { return !!sel[g.id]; });
        if (groups.length) {
          items.push({ id: it.id, kubun: it.kubun, name: it.name, groups: groups });
        }
      });
      if (items.length) {
        chapters.push({ id: ch.id, title: ch.title, name: ch.name, items: items });
      }
    });
    return chapters;
  }

  /* ---------- 共通の書式設定 ---------- */

  function style(cell, opt) {
    opt = opt || {};
    cell.font = { name: FONT, size: opt.size || FS, bold: !!opt.bold };
    cell.alignment = {
      wrapText: opt.wrap !== false,
      vertical: opt.v || 'top',
      horizontal: opt.h || 'left'
    };
    if (opt.vertical) { cell.alignment.textRotation = 'vertical'; }
    if (opt.border !== false) { cell.border = BORDER; }
    if (opt.fill) { cell.fill = HEAD_FILL; }
  }

  function pageSetup(orientation, titleRows) {
    var ps = {
      paperSize: 9,                 // A4
      orientation: orientation,
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      horizontalCentered: true,
      margins: { left: 0.4, right: 0.4, top: 0.6, bottom: 0.55, header: 0.3, footer: 0.3 }
    };
    if (titleRows) { ps.printTitlesRow = titleRows; }
    return ps;
  }

  /* ---------- 表紙・工事概要・組織表 共通 ---------- */

  // A4縦1ページのシート
  function portraitSheet(wb, name) {
    return wb.addWorksheet(name, {
      pageSetup: {
        paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 1,
        horizontalCentered: true,
        margins: { left: 0.7, right: 0.7, top: 0.8, bottom: 0.8, header: 0.3, footer: 0.3 }
      }
    });
  }

  // 左端から last 列までの列幅の合計
  function totalWidth(ws, last) {
    var w = 0;
    var i;
    for (i = 1; i <= last; i++) { w += ws.getColumn(i).width || 8; }
    return w;
  }

  // 枠なしの1行（左端から last 列まで結合）。折り返して収まらないときは行の高さを広げる
  function textLine(ws, row, last, text, size, bold, height, align) {
    ws.mergeCells(row, 1, row, last);
    var c = ws.getCell(row, 1);
    c.value = text;
    c.font = { name: FONT, size: size, bold: !!bold };
    c.alignment = { horizontal: align || 'center', vertical: 'middle', wrapText: true };
    var lines = lineCount(text, totalWidth(ws, last), size);
    ws.getRow(row).height = Math.max(height, Math.ceil(lines * size * 1.35 + 6));
  }

  // 枠付きのセル（複数セルにまたがる場合は結合する）
  function box(ws, r1, c1, r2, c2, text, opt) {
    var rr, cc;
    for (rr = r1; rr <= r2; rr++) {
      for (cc = c1; cc <= c2; cc++) { style(ws.getCell(rr, cc), opt); }
    }
    if (r1 !== r2 || c1 !== c2) { ws.mergeCells(r1, c1, r2, c2); }
    ws.getCell(r1, c1).value = text;
  }

  // 未入力の欄に入れる記号（Excel 上で書き換えて使う）
  function orMaru(v, maru) { return v ? v : (maru || '○○ ○○'); }

  /* ---------- 表紙（書籍 第Ⅲ編 2.2 総合施工計画書作成例の表紙にならう） ---------- */

  function addCover(wb, info) {
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
    ['主任監督員', '監督員', '管理技術者※', '担当者'].forEach(function (t, i) {
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
    textLine(ws, r, LAST, '　※「管理技術者」は、工事監理業務の受託者の代表となる工事監理者', 9, false, 16, 'left');
    r++;

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

  // 図の線。セルの罫線で縦線・横線を引く
  function vline(ws, r1, r2, col) {          // col 列の右側に縦線
    var rr;
    for (rr = r1; rr <= r2; rr++) {
      var c = ws.getCell(rr, col);
      var b = c.border || {};
      c.border = { top: b.top, left: b.left, bottom: b.bottom, right: THIN };
    }
  }
  function hline(ws, row, c1, c2) {          // row 行の上側に横線
    var cc;
    for (cc = c1; cc <= c2; cc++) {
      var c = ws.getCell(row, cc);
      var b = c.border || {};
      c.border = { top: THIN, left: b.left, bottom: b.bottom, right: b.right };
    }
  }

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

  // A4横1ページのシート
  function landscapeSheet(wb, name) {
    return wb.addWorksheet(name, {
      pageSetup: {
        paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 1,
        horizontalCentered: true,
        margins: { left: 0.6, right: 0.6, top: 0.7, bottom: 0.6, header: 0.3, footer: 0.3 }
      }
    });
  }

  // 表題だけのシート。中身は利用者が Excel 上で作成する（全体工程表・総合仮設計画）
  function addBlank(wb, info, name, memo) {
    var ws = landscapeSheet(wb, name);
    var cols = [];
    var i;
    for (i = 0; i < 20; i++) { cols.push({ width: 7 }); }
    ws.columns = cols;
    textLine(ws, 1, 20, name, 16, true, 30);
    textLine(ws, 2, 20, '工事名　' + (info.kojiName || ''), 11, false, 22, 'left');
    textLine(ws, 3, 20, memo, 9, false, 18, 'left');
    return ws;
  }

  /* 工種別施工計画書一覧（作成例 7.8） */
  var KOSHU = [
    '屋内配管工事施工計画書', '屋内配線工事施工計画書', '金属ダクト配線工事施工計画書',
    '分電盤、OA盤、制御盤取付け工事施工計画書', '照明器具、配線器具取付け工事施工計画書',
    '通信・情報（情報表示、拡声他）機器取付け工事施工計画書', '火災報知機器取付け工事施工計画書',
    'キュービクル式配電盤、直流電源装置搬入・据付け工事施工計画書', '自家発電装置搬入・据付け工事施工計画書',
    '太陽光発電装置搬入・据付け工事施工計画書', '接地工事施工計画書', '地中配線工事施工計画書',
    '工場立会検査計画書', '自家用電気工作物使用前自主検査計画書', '総合試運転計画書', '竣工自主検査計画書'
  ];

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

  /* ---------- 品質管理チェックシート（工種別施工計画書の別添1） ---------- */

  // シート名に使えない文字を除き、31文字以内にする
  function sheetName(s) {
    return String(s).replace(/[\[\]:*?\/\\：]/g, ' ').substring(0, 31);
  }

  function addCheckSheet(wb, info, koshuName, no, sheet) {
    var ws = portraitSheet(wb, sheetName('チェック' + no + ' ' + (sheet.item || '様式')));
    ws.columns = [{ width: 6 }, { width: 36 }, { width: 17 }, { width: 12 }, { width: 10 }, { width: 7 }, { width: 11 }];
    var H = { size: 10, v: 'middle', h: 'center', fill: true };
    var C = { size: 10, v: 'middle' };
    var CC = { size: 10, v: 'middle', h: 'center' };

    textLine(ws, 1, 7, '品質管理チェックシート', 16, true, 30);
    textLine(ws, 2, 7, '工事名　' + (info.kojiName || '') + '　／　' + koshuName, 10, false, 20, 'left');
    textLine(ws, 3, 7, sheet.title || '', 11, true, 22, 'left');

    box(ws, 4, 1, 4, 1, '項目', H);
    box(ws, 4, 2, 4, 2, sheet.item || '', C);
    box(ws, 4, 3, 4, 3, '作業場所', H);
    box(ws, 4, 4, 4, 4, '', C);
    box(ws, 4, 5, 4, 5, '作業日', H);
    box(ws, 4, 6, 4, 7, '年　月　日', CC);
    ws.getRow(4).height = Math.max(24, lineCount(sheet.item || '', 36, 10) * 14 + 8);
    ws.getRow(5).height = 8;

    box(ws, 6, 1, 7, 1, '当該\n項目', H);
    box(ws, 6, 2, 7, 2, '管理項目', H);
    box(ws, 6, 3, 7, 3, '管理基準', H);
    box(ws, 6, 4, 7, 4, '確認方法', H);
    box(ws, 6, 5, 6, 5, '担当者名', H);
    box(ws, 6, 6, 6, 7, '', C);
    box(ws, 7, 5, 7, 5, '確認日', H);
    box(ws, 7, 6, 7, 6, '合否', H);
    box(ws, 7, 7, 7, 7, '是正確認日', H);
    ws.getRow(6).height = 20;
    ws.getRow(7).height = 20;

    var rows = sheet.rows || [];
    rows.forEach(function (row, i) {
      var r = 8 + i;
      box(ws, r, 1, r, 1, '□', { size: 12, v: 'middle', h: 'center' });
      box(ws, r, 2, r, 2, row[0] || '', C);
      box(ws, r, 3, r, 3, row[1] || '', CC);
      box(ws, r, 4, r, 4, row[2] || '', CC);
      box(ws, r, 5, r, 5, '', C);
      box(ws, r, 6, r, 6, '', C);
      box(ws, r, 7, r, 7, '', C);
      var n = Math.max(lineCount(row[0] || '', 36, 10), lineCount(row[1] || '', 17, 10),
        lineCount(row[2] || '', 12, 10));
      ws.getRow(r).height = Math.max(28, n * 14 + 8);
    });
    return ws;
  }

  // 選んだ工種のチェックシートを出力する。未登録の工種は空欄の様式を1枚出力する
  function addChecks(wb, info, koshu) {
    var all = window.CHECK_DATA || {};
    if (!(koshu >= 0 && koshu <= KOSHU.length)) { koshu = 0; }
    var name = koshu < KOSHU.length ? KOSHU[koshu].replace(/施工計画書$|計画書$/, '') : (info.docTitle || '');
    var sheets = all[koshu];
    if (!sheets || !sheets.length) {
      var blank = [];
      var i;
      for (i = 0; i < 12; i++) { blank.push(['', '', '']); }
      sheets = [{ title: '', item: '', rows: blank }];
    }
    sheets.forEach(function (sh, i) { addCheckSheet(wb, info, name, i + 1, sh); });
  }

  /* ---------- 機器搬入計画書・搬入時の体制表（作成例 (8) の 5.1 ①②にならう） ---------- */

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
    var juchu = info.juchu || '○○○○電気株式会社';

    var unpan = cub ? [
      '(1) 現地構内常備のクレーン（○○ton）を使用して〔○○○○棟〕の○○階屋内：電気室へ「キュービクル・直流電源装置」を搬入・設置する。（キュービクル）は各面に分割し搬入する。',
      '(2) 変圧器のうち（動力変圧器：3φ○○kVA）×○台は取外して搬入作業を行う。',
      '(3) 機器は電気室前の〔○○○〕部分へ揚重する、機器の荷受位置から電気室に至る横引き移動通路は養生板（○○×○○×○○）を敷き並べて床面を保護養生する。',
      '(4) 「チャンネルベース」は事前搬入し（アンカーボルト）を打込据付・固定済みとする。',
      '(5) 「筐体」の据付・組立後に（組立式門形支柱）を用いて変圧器の組込作業を行う。'
    ] : ['(1) ', '(2) ', '(3) '];
    var hinmoku = cub ? [
      ['品名', 'キュービクル式配電盤、直流電源装置ほか'],
      ['個数', '〔キュービクル〕（○面型＋○面型・各面分割・変圧器取外し）×1基、（直流電源装置）一式'],
      ['①〔キュービクル〕', '○面型（○○×○○×○○）総重量：○○kg・（最大重量品：変圧器3φ○○kVA・○○kg）'],
      ['②〔直流電源装置〕', '○面型（○○×○○×○○）総重量：○○kg']
    ] : [
      ['品名', kiki ? kiki + 'ほか' : '○○○○'],
      ['個数', '○○○○ ×○基'],
      ['①〔○○○○〕', '（○○×○○×○○）総重量：○○kg']
    ];

    // [区分, 項目, 記事...]。記事は1行ずつ。3要素の配列は3列に分けて出力する
    var G = [
      ['㈠', [['目的', ['「電気設備機器（' + (kiki || '○○○○') + '）」の搬入据付工事']]]],
      ['㈡', [['搬入場所', ['○○建物：○○○室']]]],
      ['㈢実施日時', [
        ['開始', ['令和○○年○○月○○日（○）午前　○○時～']],
        ['終了', ['令和○○年○○月○○日（○）午後　○○時']],
        ['特記', ['作業時間帯：○○：○○～○○：○○']]
      ]],
      ['㈣品目', hinmoku.map(function (h) { return [h[0], [h[1]]]; })],
      ['㈤搬入業者', [
        ['住所・電話', ['○○○○○○○○　〔電話〕○○○○○○']],
        ['業者名', ['○○○○株式会社']],
        ['連絡者', ['○○○○　〔電話〕○○○○○○']]
      ]],
      ['㈥搬入方法', [
        ['運搬方法\n及び経路', unpan],
        ['車両', [
          '(1) 機器輸送車両　　　　　　　　　○○tonトラック×○台' + (cub ? '（キュービクル・変圧器）' : ''),
          '(2) 機材車両（大型機材輸送）　　　○tonユニック（○tonクレーン架装）×○台',
          '(3) 機材車両（小型・手工具輸送）　○tonトラック×○台'
        ]],
        ['体制', ['(1) 組織表　　　　別紙（搬入時の体制表）', '(2) 作業員名簿　　別紙']]
      ]],
      ['㈦安全対策', [
        ['施設', [
          '(1) 作業区域は「カラーコーン・バリケード」を設置して立入禁止とする。',
          '(2) 機械、運搬具等は必ず仮つりを行った後、本つりを行う。',
          '(3) ワイヤロープほかの使用機材の整備・点検を徹底する。',
          '(4) 作業照度不足の通路、' + (cub ? '電気室' : '搬入場所') + 'には投光器を設置し照度を確保し作業の安全に努める。'
        ]],
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
          ['クレーン・デリック運転士免許', '移動式クレーン運転士技能講習', '建設用リフト特別教育修了者'],
          ['巻上機特別教育修了者', '玉掛け技能講習修了者', '足場の組立等作業主任者'],
          ['高所作業特別教育修了者', '高所作業技能講習修了者', 'フォークリフトの運転業務']
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
    textLine(ws, r + 1, 5, '（○○ や空欄の箇所は、搬入する機器・搬入業者に合わせて書き換えてください。' +
      '該当する免許資格・選択肢は、Excel の図形（だ円）で囲んでください。）', 9, false, 30, 'left');
    return ws;
  }

  function addHannyuTaisei(wb, info) {
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
    hline(ws, 19, 7, 12);               // 2つ目の作業班へ回す線
    vline(ws, 19, 27, 12);
    hline(ws, 28, 9, 12);
    team(20, ['作業指揮者', 'クレーン\nオペレーター', '玉掛者', '合図者', '据付作業員'],
      [[1, 1, 1, 0, 1], [0, 1, 1, 0, 1], [0, 0, 0, 0, 1]]);
    team(26, ['据付作業員', '組立作業員', '監視員', '誘導員'],
      [[1, 1, 1, 1], [1, 1, 1, 0], [0, 0, 0, 0]]);
    for (r = 3; r <= 30; r++) { ws.getRow(r).height = 22; }
    ws.getRow(6).height = 12; ws.getRow(10).height = 12; ws.getRow(14).height = 12;
    ws.getRow(18).height = 12; ws.getRow(19).height = 12; ws.getRow(25).height = 12;
    ws.getRow(20).height = 28; ws.getRow(26).height = 28;
    ws.getRow(31).height = 10;
    textLine(ws, 32, LAST, '（○○○○ の箇所は、氏名・会社名・電話番号に書き換えてください。枠はセルの罫線です。）',
      9, false, 18, 'left');
    return ws;
  }

  /* ---------- 一覧表 (汎用) ---------- */

  function addTable(wb, name, title, info, heads, widths, rows, centerCols) {
    var ws = wb.addWorksheet(name, {
      pageSetup: pageSetup('landscape', '1:2'),
      views: [{ state: 'frozen', ySplit: 2 }],
      headerFooter: { oddFooter: '&C&P / &N' }
    });
    ws.columns = widths.map(function (w) { return { width: w }; });
    var n = heads.length;
    var split = Math.max(1, Math.ceil(n / 2));
    ws.mergeCells(1, 1, 1, split);
    var t = ws.getCell(1, 1);
    t.value = title;
    style(t, { size: 12, bold: true, v: 'middle', border: false, wrap: false });
    if (n > split) {
      ws.mergeCells(1, split + 1, 1, n);
      var k = ws.getCell(1, split + 1);
      k.value = info.kojiName ? ('工事名：' + info.kojiName) : '';
      style(k, { size: 9, v: 'middle', h: 'right', border: false, wrap: false });
    }
    ws.getRow(1).height = 22;
    heads.forEach(function (hd, i) {
      var c = ws.getCell(2, i + 1);
      c.value = hd;
      style(c, { size: 9, bold: true, v: 'middle', h: 'center', fill: true });
    });
    ws.getRow(2).height = 20;
    rows.forEach(function (row, ri) {
      var r = ri + 3;
      var max = 1;
      row.forEach(function (v, i) {
        var c = ws.getCell(r, i + 1);
        c.value = v;
        var center = centerCols && centerCols.indexOf(i) >= 0;
        style(c, { size: 9, v: 'middle', h: center ? 'center' : 'left' });
        max = Math.max(max, lineCount(String(v), widths[i], 9));
      });
      ws.getRow(r).height = Math.max(17, max * 12.5 + 5);
    });
    return ws;
  }

  function addItemList(wb, info, chapters) {
    var rows = [];
    var no = 0;
    chapters.forEach(function (ch) {
      ch.items.forEach(function (it) {
        no++;
        var cnt = 0;
        it.groups.forEach(function (g) { cnt += Math.max(1, g.rows.length); });
        rows.push([no, ch.title.replace(/　/g, ' '), it.kubun, it.name, it.groups.length, cnt]);
      });
    });
    return addTable(wb, '対象工事種目', '対象工事種目一覧', info,
      ['No', '工事種目', '区分', '項目', '細目数', '管理項目数'],
      [5, 26, 24, 54, 8, 11], rows, [0, 4, 5]);
  }

  function addDocList(wb, info, chapters) {
    var rows = [];
    chapters.forEach(function (ch) {
      ch.items.forEach(function (it) {
        it.groups.forEach(function (g) {
          g.rows.forEach(function (row) {
            parseDocs(row[4]).forEach(function (doc) {
              rows.push([ch.name, it.name, g.e, doc, row[3], row[6], '', '']);
            });
          });
        });
      });
    });
    return addTable(wb, '作成・提出資料一覧', '作成・提出資料一覧', info,
      ['工事種目', '項目', '細目', '資料名', '承・検・立', '提出', '作成済', '備考'],
      [20, 26, 24, 44, 9, 6, 7, 16], rows, [4, 5, 6]);
  }

  function addPhotoList(wb, info, chapters) {
    var rows = [];
    chapters.forEach(function (ch) {
      ch.items.forEach(function (it) {
        it.groups.forEach(function (g) {
          g.rows.forEach(function (row) {
            if (row[5]) {
              rows.push([ch.name, it.name, g.e, row[5].replace(/\n/g, ''), row[7], '', '']);
            }
          });
        });
      });
    });
    return addTable(wb, '写真撮影項目一覧', '写真撮影項目一覧', info,
      ['工事種目', '項目', '細目', '撮影対象', '提出', '撮影済', '備考'],
      [20, 28, 26, 46, 6, 7, 18], rows, [4, 5]);
  }

  /* ---------- 管理表 (章ごと) ---------- */

  function addChapter(wb, info, ch, withPage) {
    var ncol = withPage ? 14 : 13;
    var ws = wb.addWorksheet(ch.title.replace(/　/g, ' '), {
      pageSetup: pageSetup('landscape', '1:3'),
      views: [{ state: 'frozen', ySplit: 3 }],
      headerFooter: { oddFooter: '&C&P / &N' }
    });
    ws.columns = W.slice(0, ncol).map(function (w) { return { width: w }; });

    // 1行目: タイトルと工事名
    ws.mergeCells(1, 1, 1, 7);
    var t = ws.getCell(1, 1);
    t.value = ch.title + '　管理表';
    style(t, { size: 11, bold: true, v: 'middle', border: false, wrap: false });
    ws.mergeCells(1, 8, 1, ncol);
    var k = ws.getCell(1, 8);
    k.value = info.kojiName ? ('工事名：' + info.kojiName) : '';
    style(k, { size: 9, v: 'middle', h: 'right', border: false, wrap: false });
    ws.getRow(1).height = 20;

    // 2〜3行目: 見出し (元の管理表と同じ構成)
    var r2 = { 1: '管理内容', 7: '管理値・\n管理基準', 8: '管理方法', 10: '作成資料', 12: '提出資料', 14: '掲載\nページ' };
    var r3 = { 1: '区分', 2: '項目', 3: '標仕', 4: '標準図', 5: '細目', 6: '要点', 10: '資料', 11: '写真', 12: '資料', 13: '写真' };
    var c;
    for (c = 1; c <= ncol; c++) {
      var a = ws.getCell(2, c);
      var b = ws.getCell(3, c);
      if (r2[c]) { a.value = r2[c]; }
      if (r3[c]) { b.value = r3[c]; }
      style(a, { v: 'middle', h: 'center', fill: true, bold: true });
      style(b, { v: 'middle', h: 'center', fill: true, bold: true });
    }
    ws.mergeCells(2, 1, 2, 6);
    ws.mergeCells(2, 7, 3, 7);
    ws.mergeCells(2, 8, 3, 9);
    ws.mergeCells(2, 10, 2, 11);
    ws.mergeCells(2, 12, 2, 13);
    if (withPage) { ws.mergeCells(2, 14, 3, 14); }
    ws.getRow(2).height = 14;
    ws.getRow(3).height = 24;

    // 本体
    var heights = {};       // 行番号 → 高さ
    var blocks = [];        // 細目ごとの {kubun, item, name, r1, r2}
    var r = 4;
    ch.items.forEach(function (it) {
      it.groups.forEach(function (g) {
        var rows = g.rows.length ? g.rows : [['', '', '', '', '', '', '', '', '']];
        var start = r;
        rows.forEach(function (row) {
          var max = 1;
          for (var i = 0; i < ncol - 5; i++) {
            var cell = ws.getCell(r, 6 + i);
            var v = row[i] || '';
            if (v) { cell.value = v; }
            var center = (i === 3 || i === 6 || i === 7 || i === 8);
            style(cell, { h: center ? 'center' : 'left' });
            max = Math.max(max, lineCount(v, W[5 + i]));
          }
          for (var j = 1; j <= 5; j++) { style(ws.getCell(r, j)); }
          heights[r] = heightFor(max);
          r++;
        });
        var end = r - 1;
        // 標仕・標準図・細目は細目単位で縦結合
        ws.getCell(start, 3).value = g.c || null;
        ws.getCell(start, 4).value = g.d || null;
        ws.getCell(start, 5).value = g.e || null;
        if (end > start) {
          ws.mergeCells(start, 3, end, 3);
          ws.mergeCells(start, 4, end, 4);
          ws.mergeCells(start, 5, end, 5);
        }
        var need = heightFor(Math.max(lineCount(g.c, W[2]), lineCount(g.d, W[3]), lineCount(g.e, W[4])));
        pad(heights, start, end, need);
        blocks.push({ kubun: it.kubun, item: it.id, name: it.name, r1: start, r2: end });
      });
    });

    // 項目(B列)・区分(A列)を結合。長い場合は CHUNK ごとに区切って見出しを繰り返す
    mergeLabels(ws, heights, blocks, 2, function (b) { return b.item; },
      function (b) { return b.name; },
      function (text) { return heightFor(lineCount(text, W[1])); }, false);
    mergeLabels(ws, heights, blocks, 1, function (b) { return b.kubun; },
      function (b) { return b.kubun; },
      function (text) { return text.length * (FS + 2.5) + 8; }, true);

    Object.keys(heights).forEach(function (key) {
      ws.getRow(Number(key)).height = Math.round(heights[key] * 10) / 10;
    });
    return ws;
  }

  // r1〜r2 の合計高さが need に満たなければ最終行を広げる
  function pad(heights, r1, r2, need) {
    var sum = 0;
    for (var i = r1; i <= r2; i++) { sum += heights[i]; }
    if (sum < need) { heights[r2] += (need - sum); }
  }

  function mergeLabels(ws, heights, blocks, col, keyFn, textFn, needFn, vertical) {
    var i = 0;
    while (i < blocks.length) {
      var key = keyFn(blocks[i]);
      var start = blocks[i].r1;
      var end = blocks[i].r2;
      var acc = 0;
      var j = i;
      while (j < blocks.length && keyFn(blocks[j]) === key) {
        for (var q = blocks[j].r1; q <= blocks[j].r2; q++) { acc += heights[q]; }
        end = blocks[j].r2;
        j++;
        if (acc >= CHUNK) { break; }
      }
      var text = textFn(blocks[i]);
      var cell = ws.getCell(start, col);
      cell.value = text;
      if (vertical) {
        style(cell, { vertical: true, h: 'center', v: 'top', wrap: false });
      }
      if (end > start) { ws.mergeCells(start, col, end, col); }
      pad(heights, start, end, needFn(text));
      i = j;
    }
  }

  /* ---------- ブック全体 ---------- */

  function build(data, state) {
    var ExcelJS = window.ExcelJS;
    if (!ExcelJS) { return Promise.reject(new Error('ExcelJS が読み込まれていません')); }
    var info = state.info || {};
    var opt = state.opt || {};
    var chapters = collect(data, state.sel || {});
    // 管理表を出力しない書類（総合施工計画書）は、工事種目を選んでいなくても作成できる
    if (!chapters.length && opt.kanri !== false) {
      return Promise.reject(new Error('工事種目が選択されていません'));
    }

    var wb = new ExcelJS.Workbook();
    wb.creator = info.juchu || 'sekokan';
    wb.created = new Date();

    if (opt.cover !== false) { addCover(wb, info); }
    if (opt.gaiyo !== false) { addGaiyo(wb, info, chapters, opt.kanri !== false); }
    if (opt.soshiki !== false) { addSoshiki(wb, info); addHoan(wb); addHinshitsu(wb, info); }
    if (opt.tenpu !== false) {
      addBlank(wb, info, '全体工程表', '（このシートに全体工程表を作成してください。）');
      addKoshu(wb);
      addAnzenSoshiki(wb, info);
      addAnzenKeikaku(wb, info);
      addRenraku(wb);
      addBlank(wb, info, '総合仮設計画', '（このシートに総合仮設計画図を作成してください。）');
    }
    if (chapters.length) {
      if (opt.list !== false) { addItemList(wb, info, chapters); }
      if (opt.kanri !== false) {
        chapters.forEach(function (ch) { addChapter(wb, info, ch, opt.page !== false); });
      }
      if (opt.docs !== false) { addDocList(wb, info, chapters); }
      if (opt.photos !== false) { addPhotoList(wb, info, chapters); }
    }
    if (opt.hannyu === true) { addHannyu(wb, info, Number(state.koshu) || 0); addHannyuTaisei(wb, info); }
    if (opt.check === true) { addChecks(wb, info, Number(state.koshu) || 0); }
    if (!wb.worksheets.length) { return Promise.reject(new Error('出力するシートがありません。出力内容を選んでください')); }

    return wb.xlsx.writeBuffer().then(function (buf) {
      return new Blob([buf], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
    });
  }

  window.SekokanExport = {
    build: build,
    collect: collect,
    parseDocs: parseDocs,
    wareki: wareki,
    KOSHU: KOSHU
  };
})();
