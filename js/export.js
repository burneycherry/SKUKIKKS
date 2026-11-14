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

  /* ---------- 一覧表 (汎用) ---------- */

  /* ---------- 使用前自主検査の検査項目（作成例 (14) の 6.1 ⑥ 自主検査の方法及び判定基準。検査項目ごとに1シート） ---------- */

  var KENSA_CHUI = '使用前には必ず次の三つのチェックを行う。\n' +
    '①電池チェック：スイッチを押さずにL端子リード線を右側面の電池チェック端子に触れ、指針が目盛板上のBマーク内にあること。\n' +
    '②零チェック　：L端子とE端子を短絡し、スイッチを入れて指針が零を示すことを確かめる。\n' +
    '③開放チェック：L端子とE端子を開放し、スイッチを入れて指針が∞を示すことを確かめる。';

  // { no, title, mokuteki, naiyo, kijun, chui, biko }
  var KENSA = [
    { no: '①', title: '外観検査',
      mokuteki: '検査対象となる電気工作物が届出をした工事計画に従って行われたものであること、及びその電気工作物が電気設備に関する技術基準に適合しているものであることを確認する。',
      naiyo: '検査対象となる電気工作物の設置状況について、工事計画に従って工事が行われていること、及び電気設備に関する技術基準に適合していることを目視により確認する。\n' +
        'なお、判定基準の①、②、③、⑩、⑪を確認する場合は、書類等によって確認することもできる。',
      kijun: '①必要な箇所に所定の接地が行われていること（電技解釈第17条～19条、第21条、第24条、第28条、第29条、第37条）\n' +
        '②アークを発生する器具と可燃性物質との離隔が十分であること（電技解釈第23条）\n' +
        '③高圧用の機械器具の充電部が、取扱者が容易に触れないように施設されていること（電技解釈第21条、第22条）\n' +
        '④高圧の電路において電線及び電気機械器具を保護するため必要な箇所に過電流遮断器が施設されていること（電技解釈第34条、第35条）\n' +
        '⑤高圧電路中の過電流遮断器の開閉状態が容易に確認できること（電技解釈第34条）\n' +
        '⑥高圧の電路に地絡を生じたときに自動的に電路を遮断する装置が必要な箇所に施設されていること（電技解釈第36条）\n' +
        '⑦高圧の電路において、架空電線の引込口及び引出口又はこれに近接する箇所に避雷器が施設されていること（電技解釈第37条）\n' +
        '⑧変電所、開閉所若しくはこれらに準ずる場所（以下「変電所等に準ずる場所」という）の周囲に、さく、へい等が施設されており、出入口に施錠装置及び立入禁止表示が施設されていること（電技解釈第38条）\n' +
        '⑨中性点直接接地式電路に接続する変圧器には、油流出防止設備が施設されていること（電技第19条第10項）\n' +
        '⑩ガス絶縁機器等の圧力容器が規定どおり施設されていること（電技解釈第40条）\n' +
        '⑪検査の対象となる電気工作物が工事計画書の記載事項どおりに施設されていること',
      chui: '',
      biko: '1. 検査記録は「外観検査成績表」に結果を記入する。\n2. 対象外については「外観検査成績表」の判定欄に「該当外」を付すこと。' },
    { no: '②-イ', title: '接地抵抗測定',
      mokuteki: '異常時の電位上昇、高電圧の侵入等による感電、火災その他人体に危害を及ぼし、又は物件への損傷を与える事がないよう、各種接地工事の接地抵抗値が電気設備の技術基準の解釈に示す値以下であることを確認する。',
      naiyo: '(1)接地抵抗計による測定\n' +
        '　設置面積の狭い単独極に適用され、測定接地極と補助極（P極、C極）とほぼ一直線上間隔（10m以上）になるような補助極を選定あるいは設け測定する。\n\n' +
        '(2)測定手順\n' +
        '①接地端子盤のP端子に接地抵抗計のP端子（電圧電極）、C端子に接地抵抗計のC端子（電流電極）にリード線をそれぞれ接続する。\n' +
        '②接地抵抗計のE端子（接地端子）のリード線を測定しようとする極に接続する。\n' +
        '③接地抵抗計の切替スイッチをBにし、測定スイッチを押し指針が「CHECK」の枠内にあることを確認する。（バッテリー確認）\n' +
        '④切替スイッチをVにして、地電圧を測定する。等価地電圧は5V以下で測定する。切替スイッチをΩレンジにして、検流計のバランスをとりながらダイヤルを調整する。検流値が零を示した時のダイヤル指示値が測定値となる。',
      kijun: '電気設備の技術基準の解釈（第17条）による接地工事の種類と接地抵抗値\n\n' +
        '1. A種接地工事\n' +
        '　①電圧の種別による機器\n　　高圧用機械器具の金属製外箱等\n' +
        '　②接地抵抗許容値\n　　10Ω以下\n' +
        '　③接地線の種類\n　　引張強さ1.04kN以上の金属線又は直径2.6mm以上の軟銅線\n\n' +
        '2. B種接地工事\n' +
        '　①高圧用電路と低圧電路とを結合する変圧器の低圧側中性点\n' +
        '　②接地抵抗許容値\n　　計算値以下　※「B種接地抵抗許容値の算出」のシート参照\n' +
        '　③接地線の種類\n　　引張強さ2.46kN以上の金属線又は直径4mm以上の軟銅線\n\n' +
        '3. C種接地工事\n' +
        '　①電圧の種別による機器\n　　低圧用機械器具の金属製外箱等（300Vを超える）\n' +
        '　②接地抵抗許容値\n　　10Ω以下\n' +
        '　③接地線の種類\n　　引張強さ0.39kN以上の金属線又は直径1.6mm以上の軟銅線\n\n' +
        '4. D種接地工事\n' +
        '　①電圧の種別による機器\n　　低圧用機械器具の金属製外箱等（300V以下）\n　　高圧計器用変成器の2次側電路\n' +
        '　②接地抵抗許容値\n　　100Ω以下\n' +
        '　③接地線の種類\n　　引張強さ0.39kN以上の金属線又は直径1.6mm以上の軟銅線',
      chui: '',
      biko: '1. 測定記録は「接地抵抗測定成績表」に結果を記入する。\n2. 構造体接地については、事前に大地比抵抗率測定を行った記録にて、基準値内（2Ω以下）であることを確認しておく。' },
    { no: '②-ロ', title: 'B種接地抵抗許容値の算出', calc: true },
    { no: '②-ハ', title: '接地導通試験',
      mokuteki: '各受配電機器に接地工事の種類に応じて確実に配線されていることを確認する。',
      naiyo: '(1)接地端子盤から各受配電盤機器への接地線の導通確認を行う。\n\n' +
        '(2)接地端子盤の二次側を切り離した後、予備線を使用し事前に機能チェックを行ったブザーにて各接地配線の導通があることを確認する。',
      kijun: '接地極と受配電盤機器との間に導通があり、ブザーが鳴動すること。',
      zu: '【参考図】（接地端子盤・二次配線切離し・受配電機器・接地母線・予備線・ブザーの図は、この欄に作成してください。）',
      chui: '接地端子の復旧は確実に行うこと。',
      biko: '試験記録は「接地導通試験成績表」に結果を記入する。' },
    { no: '③-1', title: '絶縁抵抗測定／（高圧）',
      mokuteki: '高圧電路の絶縁状態の確認は、電気設備技術基準では絶縁耐力試験で判定することになっている。その予備試験として絶縁抵抗を測定して、その値から絶縁耐力試験を実施することが適当か否かを判定する。また、絶縁耐力試験後にも絶縁抵抗を測定し、試験前と比較する。',
      naiyo: '(1)測定機器は電池式絶縁抵抗計（JIS C 1302）によるものとし、その定格電圧は1,000Vを選定する。\n\n' +
        '(2)測定手順\n' +
        '①測定端子の接地側（E）端子をアース端子に、線路側（L）端子を被測定回路（電路又は機器）に接続し測定する。\n' +
        '②測定値は原則として「1分値」を採用する。ただし、測定値が変化している場合は安定するまで測定時間を延長し、最大「3分値」の数値をもって採用する。\n' +
        '③30秒以内で2,000MΩ以上になる場合は、その測定時間を採用する。',
      kijun: '大地と絶縁されていること。絶縁とは以下の数値基準に適合している場合とする。\n\n' +
        '　高圧電路と大地間　30MΩ以上\n　（工場立会検査計画書の 3.1「キュービクル式配電盤」参照）',
      chui: KENSA_CHUI,
      biko: '測定記録は「交流絶縁耐力測定表」に結果を記入する。' },
    { no: '③-2', title: '絶縁抵抗測定／（低圧）',
      mokuteki: '電気機器及び電気回路の絶縁状態の良否を判断するために実施する。',
      naiyo: '(1)測定機器は電池式絶縁抵抗計（JIS C 1302）によるものとし、その定格電圧は500Vを選定する。\n\n' +
        '※制御機器等が接続されている場合で、絶縁抵抗測定により制御機器等の損傷が予想される場合の定格測定電圧は次による。\n' +
        '　電路の使用電圧　定格測定電圧\n' +
        '　　100V級　　　　125〔V〕\n' +
        '　　200V級　　　　250〔V〕\n' +
        '　　400V級　　　　500〔V〕\n\n' +
        '(2)測定手順\n' +
        '①絶縁抵抗計の電圧で損傷するおそれのある機器は、事前に測定回路から切り離す。\n' +
        '②測定端子の接地側（E）端子をアース端子に、機器側（L）端子を被測定回路に接続し測定する（電路と大地間の測定）。\n' +
        '③測定端子の接地側（E）端子、線路側（L）端子をそれぞれ被測定回路に接続し測定する（電線相互間の測定）。\n' +
        '④絶縁抵抗計のテストボタンはE端子、L端子を接続した後押す。\n' +
        '⑤対地静電容量の大きな電路では、測定値が安定するまで時間がかかるので、測定値が安定した時点の値を測定値とする。\n' +
        '⑥測定後は電路に直流電荷が残留していることがあるので電荷の放電を行う。',
      kijun: '低圧電路と大地間、電線相互間　5MΩ以上\nただし、機器が接続された状態では1MΩ以上\n\n\n' +
        '参考：電気設備の技術基準（第58条）では、次のように規定している。\n' +
        '（電路の使用電圧の区分に応じ下記の値以上とする）\n' +
        '　300V以下（対地電圧が150V以下の場合）　0.1MΩ\n' +
        '　300V以下（その他の場合）　　　　　　　0.2MΩ\n' +
        '　300Vを超えるもの　　　　　　　　　　　0.4MΩ',
      chui: KENSA_CHUI,
      biko: '測定記録は「低圧幹線絶縁抵抗測定表」及び「低圧母線絶縁抵抗測定表」に結果を記入する。' },
    { no: '④-1', title: '絶縁耐力試験',
      mokuteki: '絶縁耐力試験は電気工作物の運転開始に先立ち電路や機器の絶縁信頼度を判定するために電気設備の技術基準の解釈（第15・16条）に基づき、高圧機器を含む主回路と大地間へ規定された試験電圧を規定時間印加し、絶縁に異常のないことを確認する。',
      naiyo: '下記に示す試験電圧を電路と大地間に連続して10分間印加し、これに耐えるかを確認する。\n\n' +
        '・交流の場合\n　高圧（公称電圧6,600V）\n　6,600×1.15/1.1×1.5＝10,350（V）',
      kijun: '電気設備の技術基準の解釈\n' +
        '　第15条　高圧又は特別高圧の電路の絶縁性能\n' +
        '　第16条　機械器具等の電路の絶縁性能\n\n' +
        '試験電圧を連続して10分間加えた後、絶縁抵抗測定を行い絶縁に異常のないこと。\n' +
        'なお、中間計測記録は1分値、5分値、9分値とする。\n\n' +
        '※第16条第1項第2号、第16条第6項第3号における「電路の絶縁耐力の確認方法」において、日本電気技術規格委員会規格JESC E7001（2025）に基づく確認方法（常規対地電圧を印加する方法）を採用する場合は、機器がJEC、JISで定める商用周波耐電圧試験による絶縁耐電圧試験に合格したもので、設置場所においてもその性能が維持されると判断できる場合のみ採用可能とする。\n' +
        '常規対地電圧を連続して10分間加え、絶縁に異常のないこと。',
      chui: '',
      biko: '試験記録は「交流絶縁耐力測定表」に結果を記入する。' }
  ];

  // [作業手順, 作業の要点, 会場外への指示と指示受け者の有無]
  var KENSA_TEJUN = [
    ['①試験回路説明', '・現地掲示物を用いて試験体制、試験方法、試験回路等を検査責任者へ説明する。', 0],
    ['②試験開始合図', '・係員全員へ試験を開始する旨連絡する。', 1],
    ['③監視係配置確認', '・監視係が所定の位置へ就いていることを確認する。', 1],
    ['④安全確認', '・試験範囲の安全確認を行う旨を指示し、安全であることを確認する。', 1],
    ['⑤絶縁抵抗測定', '・検査責任者が計器の読みを確認する。', 1],
    ['⑥放電接地', '・絶縁抵抗測定点にて行う。', 0],
    ['⑦印加開始', '・試験電源を投入し、電圧を印加する（1/2試験電圧）。\n・充電電流、電流レンジの確認。', 1],
    ['⑧検電', '・試験責任者は各所へ検電確認を指示する。\n・各所検電確認を行い、全ての試験回路において有電圧を確認する。\n・検電確認終了の旨を試験責任者へ連絡する。', 1],
    ['⑨試験電圧印加', '・試験電圧まで昇圧する。\n・ストップウォッチをスタートさせ、主任技術者が印加開始時刻を確認する。\n・測定時は検査責任者が電圧計、電流計の読みを確認する。', 1],
    ['⑩降圧・電源断', '・降圧前に検査責任者の了解を得る。', 1],
    ['⑪無電圧確認', '・印加点にて検電し無電圧であることを確認する。', 0],
    ['⑫放電接地', '・印加点にて行う。', 0],
    ['⑬絶縁抵抗測定', '・検査責任者が計器の読みを確認する。', 1],
    ['⑭放電接地', '・絶縁抵抗測定点にて行う。', 0],
    ['⑮試験終了', '・試験終了の旨を係員全員へ連絡する。', 1]
  ];

  // 様式の上段（検査項目・No.・表題）。左の見出しは1〜split列、表題は残りの列
  function kensaHead(ws, no, title, split, last) {
    var H = { size: 10, v: 'middle', h: 'center', fill: true };
    box(ws, 1, 1, 1, split, '検査項目', H);
    box(ws, 2, 1, 2, 1, 'No.', H);
    box(ws, 2, 2, 2, split, no, { size: 10, v: 'middle', h: 'center' });
    box(ws, 1, split + 1, 2, last, title, { size: 13, bold: true, v: 'middle', h: 'center' });
    ws.getRow(1).height = 20;
    ws.getRow(2).height = 20;
  }

  function kensaHeight(text, width) {
    return Math.max(22, lineCount(text, width, 10) * 13.5 + 10);
  }

  function addKensaSheet(wb, k) {
    var ws = portraitSheet(wb, sheetName(k.no + ' ' + k.title.replace('／', '')));
    ws.columns = [{ width: 6 }, { width: 8 }, { width: 26 }, { width: 52 }];
    var H = { size: 10, v: 'middle', h: 'center', fill: true };
    var T = { size: 10, v: 'top' };
    var M = { size: 10, v: 'middle' };
    kensaHead(ws, k.no, k.title, 2, 4);
    box(ws, 3, 1, 3, 1, '目的', H);
    box(ws, 3, 2, 3, 4, k.mokuteki, M);
    ws.getRow(3).height = kensaHeight(k.mokuteki, 86);
    box(ws, 4, 1, 4, 3, '検査内容', H);
    box(ws, 4, 4, 4, 4, '判定基準', H);
    ws.getRow(4).height = 20;
    box(ws, 5, 1, 5, 3, k.naiyo, T);
    box(ws, 5, 4, 5, 4, k.kijun, T);
    ws.getRow(5).height = Math.max(k.zu ? 110 : 330, kensaHeight(k.naiyo, 40), kensaHeight(k.kijun, 52));
    var r = 6;
    if (k.zu) {
      box(ws, r, 1, r, 4, k.zu, T);
      ws.getRow(r).height = 240;
      r++;
    }
    box(ws, r, 1, r, 1, '注意\n事項', H);
    box(ws, r, 2, r, 4, k.chui, M);
    ws.getRow(r).height = Math.max(40, kensaHeight(k.chui, 86));
    box(ws, r + 1, 1, r + 1, 1, '備考', H);
    box(ws, r + 1, 2, r + 1, 4, k.biko, M);
    ws.getRow(r + 1).height = Math.max(36, kensaHeight(k.biko, 86));
    return ws;
  }

  // B種接地抵抗許容値の算出（色付きの欄を書き換えると Excel 上で再計算される）
  function addKensaBshu(wb, k) {
    var ws = portraitSheet(wb, sheetName(k.no + ' ' + k.title));
    ws.columns = [{ width: 6 }, { width: 8 }, { width: 44 }, { width: 34 }];
    var L = { size: 10, v: 'middle' };
    var C = { size: 10, v: 'middle', h: 'center' };
    var INPUT = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
    kensaHead(ws, k.no, k.title, 2, 4);
    ws.getRow(3).height = 8;
    var r = 4;
    function line(t, bold) { textLine(ws, r, 4, t, bold ? 11 : 10, bold, 20, 'left'); r++; }
    line('電技解釈第17条による計算式', true);
    line('（電線にケーブル以外のものを使用する電路と電線にケーブルを使用する電路とからなる電路）');
    line('　I1 ＝ 1 ＋ (V/3×L－100)／150 ＋ (V/3×L\'－1)／2');
    line('　右辺の第2項及び第3項の値は、それぞれの値が負となる場合は、0とする。');
    line('　I1 の値は、小数点以下は切り上げとする。I1 が2未満となる場合は、2とする。');
    line('　　I1：1線地絡電流［A］　　V：電路の公称電圧を1.1で除した電圧［kV］');
    line('　　L：高圧ケーブルを除いた電線延長［km］　　L\'：高圧ケーブル線路延長［km］');
    line('B種接地抵抗許容値', true);
    line('　R ＝ 150／I1［Ω］　　R：B種接地抵抗値［Ω］　　I1：1線地絡電流［A］');
    ws.getRow(r).height = 10; r++;
    line('当事業場のB種接地抵抗許容値の算出（色付きの欄を書き換えると再計算されます）', true);
    var r0 = r;
    // [項目, 値または計算式, 計算結果の初期値, 入力欄か]
    var rows = [
      ['電路の公称電圧［kV］', 6.6, null, true],
      ['高圧ケーブルを除いた電線延長 L［km］', 0, null, true],
      ['高圧ケーブル線路延長 L\'［km］', 3.584, null, true],
      ['V（公称電圧を1.1で除した電圧）［kV］', 'ROUND(D' + r0 + '/1.1,3)', 6, false],
      ['第2項 (V/3×L－100)／150（負のときは0）', 'MAX(0,(D' + (r0 + 3) + '/3*D' + (r0 + 1) + '-100)/150)', 0, false],
      ['第3項 (V/3×L\'－1)／2（負のときは0）', 'MAX(0,(D' + (r0 + 3) + '/3*D' + (r0 + 2) + '-1)/2)', 3.084, false],
      ['I1 の計算値［A］', '1+D' + (r0 + 4) + '+D' + (r0 + 5), 4.084, false],
      ['I1（小数点以下切り上げ。2未満は2）［A］', 'MAX(2,ROUNDUP(D' + (r0 + 6) + ',0))', 5, false],
      ['B種接地抵抗許容値 R ＝ 150／I1［Ω］', '150/D' + (r0 + 7), 30, false]
    ];
    rows.forEach(function (row) {
      box(ws, r, 1, r, 3, row[0], L);
      box(ws, r, 4, r, 4, '', C);
      if (row[3]) {
        ws.getCell(r, 4).value = row[1];
        ws.getCell(r, 4).fill = INPUT;
      } else {
        ws.getCell(r, 4).value = { formula: row[1], result: row[2] };
      }
      ws.getRow(r).height = 22;
      r++;
    });
    ws.getRow(r).height = 10; r++;
    ws.mergeCells(r, 1, r, 4);
    var c = ws.getCell(r, 1);
    c.value = { formula: '"※当事業場のB種接地抵抗許容値は "&D' + (r0 + 8) + '&"［Ω］以下とする。"',
      result: '※当事業場のB種接地抵抗許容値は 30［Ω］以下とする。' };
    c.font = { name: FONT, size: 11, bold: true };
    c.alignment = { horizontal: 'right', vertical: 'middle' };
    ws.getRow(r).height = 24;
    return ws;
  }

  // 絶縁耐力試験／手順（表）
  function addKensaTejun(wb) {
    var ws = portraitSheet(wb, sheetName('④-2 絶縁耐力試験 手順'));
    ws.columns = [{ width: 7 }, { width: 12 }, { width: 40 }, { width: 17 }, { width: 12 }, { width: 9 }];
    var H = { size: 10, v: 'middle', h: 'center', fill: true };
    var C = { size: 10, v: 'middle', h: 'center' };
    var M = { size: 10, v: 'middle' };
    kensaHead(ws, '④-2', '絶縁耐力試験／手順', 2, 6);
    textLine(ws, 3, 6, '(1) 絶縁耐力試験手順', 10, false, 22, 'left');
    box(ws, 4, 1, 4, 2, '作業手順', H);
    box(ws, 4, 3, 4, 3, '作業の要点', H);
    box(ws, 4, 4, 4, 4, '会場外への指示', H);
    box(ws, 4, 5, 4, 5, '指示受け者', H);
    box(ws, 4, 6, 4, 6, 'チェック', H);
    ws.getRow(4).height = 20;
    KENSA_TEJUN.forEach(function (t, i) {
      var r = 5 + i;
      box(ws, r, 1, r, 2, t[0], M);
      box(ws, r, 3, r, 3, t[1], M);
      box(ws, r, 4, r, 4, t[2] ? '連絡係' : '', C);
      box(ws, r, 5, r, 5, t[2] ? '監視係' : '', C);
      box(ws, r, 6, r, 6, '', C);
      ws.getRow(r).height = Math.max(24, lineCount(t[1], 40, 10) * 13.5 + 10, lineCount(t[0], 19, 10) * 13.5 + 10);
    });
    return ws;
  }

  // 絶縁耐力試験／手順（フロー）。枠はセルの罫線、矢印は文字で描く
  function addKensaFlow(wb) {
    var ws = portraitSheet(wb, sheetName('④-3 絶縁耐力試験 手順フロー'));
    ws.columns = [{ width: 9 }, { width: 9 }, { width: 7 }, { width: 9 }, { width: 9 }, { width: 9 }, { width: 4 },
      { width: 8 }, { width: 8 }, { width: 8 }, { width: 18 }];
    var B = { size: 11, v: 'middle', h: 'center' };
    var Q = { size: 11, v: 'middle', h: 'center', fill: true };   // 判定（原本はだ円）
    var r;
    function txt(row, c1, c2, text, align, size) {
      if (c1 !== c2) { ws.mergeCells(row, c1, row, c2); }
      var c = ws.getCell(row, c1);
      c.value = text;
      c.font = { name: FONT, size: size || 10 };
      c.alignment = { horizontal: align || 'center', vertical: 'middle' };
    }
    kensaHead(ws, '④-3', '絶縁耐力試験／手順（フロー）', 2, 11);
    textLine(ws, 3, 11, '(2) 絶縁耐力試験手順フロー', 10, false, 22, 'left');
    ws.getRow(4).height = 8;
    for (r = 5; r <= 23; r++) { ws.getRow(r).height = r % 2 ? 26 : 20; }

    // 中央の流れ（5行目から1つおき）
    var mid = [
      ['①試験回路説明', B, '↓'], ['②試験開始合図', B, '↓'], ['③監視係配置', B, '↓'], ['④安全確認', B, '↓'],
      ['⑤絶縁抵抗測定', Q, '↓ YES'], ['⑥放電接地', B, '↓'], ['⑦印加開始', B, '↓ 1/2試験電圧まで昇圧'],
      ['異常の発生', Q, '↓ NO'], ['⑧検電', Q, '↓ YES'], ['Ⓐ', B, '']
    ];
    mid.forEach(function (m, i) {
      var row = 5 + i * 2;
      if (i === 9) { box(ws, row, 5, row, 5, m[0], B); } else { box(ws, row, 4, row, 6, m[0], m[1]); }
      if (m[2]) { txt(row + 1, 4, 6, m[2]); }
    });

    // 左の流れ（異常時。下から上へ進む）
    txt(5, 1, 3, '（①試験回路説明へ）→', 'right');
    txt(6, 1, 2, '↑');
    [['印加範囲確認', 7], ['問題箇所改修', 9], ['放電接地', 11], ['無電圧確認', 13], ['非常停止', 19]].forEach(function (b) {
      box(ws, b[1], 1, b[1], 2, b[0], B);
    });
    [8, 10, 12, 14, 15, 16, 17, 18].forEach(function (row) { txt(row, 1, 2, '↑'); });
    txt(13, 3, 3, '← NO');
    txt(19, 3, 3, '← YES');
    txt(20, 1, 2, '↑');
    txt(21, 1, 2, '↑');
    txt(21, 3, 3, '← NO');

    // 右の流れ（Ⓐ の続き）
    var right = [
      ['Ⓐ', '↓', '試験電圧まで昇圧'], ['⑨試験電圧印加', '↓', ''], ['印加範囲確認', '↓', '主任技術者は計測時間'],
      ['1、5、9分値測定', '↓', ''], ['⑩降圧・電源断', '↓', '降圧する前に主任技術者の了承を得る。'],
      ['⑪無電圧確認', '↓', '検電'], ['⑫放電接地', '↓', ''], ['⑬絶縁抵抗測定', '↓', ''], ['⑭放電接地', '↓', ''],
      ['⑮試験終了', '', '']
    ];
    right.forEach(function (m, i) {
      var row = 5 + i * 2;
      if (i === 0) { box(ws, row, 9, row, 9, m[0], B); } else { box(ws, row, 8, row, 10, m[0], B); }
      if (m[1]) { txt(row + 1, 8, 10, m[1]); }
      if (m[2]) {
        var c = ws.getCell(row + 1, 11);
        c.value = m[2];
        c.font = { name: FONT, size: 9 };
        c.alignment = { vertical: 'middle', wrapText: true };
      }
    });
    txt(7, 11, 11, '（10分間の印加）', 'left', 9);
    ws.getRow(14).height = 26;
    ws.getRow(24).height = 10;
    textLine(ws, 25, 11, '（色付きの枠は判定です。YES・NO の先へ進みます。左の列は異常時の流れで、下から上へ進み ① に戻ります。）',
      9, false, 18, 'left');
    return ws;
  }

  // 絶縁耐力試験／安全対策及び注意事項（本文だけの様式）
  var KENSA_ANZEN = '1. 各試験は検査責任者の指示に従い実施すること。\n' +
    '2. TBM（※1）・KYM（※2）を実施し、関係者に試験手順、印加範囲、開始時間等を十分に把握させ危険ポイントを指示する。\n' +
    '3. 試験器及び充電部のまわりはバリケード（安全ネット等）により危険表示を行い、試験中パトライト（回転灯）を点灯し、監視係を配置して関係者以外の立入を禁止する。\n' +
    '4. 被試験場所には試験中である旨の表示をし、かつ、安全監視員を配置して、関係者以外の立入を禁止する。\n' +
    '5. 各試験場所にトランシーバ等を設置し、指示・連絡・合図等、確認を徹底する。\n' +
    '6. 監視係は異音、異臭等、異常が発生した場合は、試験責任者に速やかに報告する。\n' +
    '7. 試験器の操作場所には絶縁ゴムマットを設置する。\n' +
    '8. 検電係は安全帽、絶縁ゴム手袋を着用する。\n\n' +
    '※1　TBM……ツール・ボックス・ミーティングの略\n' +
    '　　　その日の作業内容、注意点を作業責任者を中心に話し合う短時間で行うミーティング\n' +
    '※2　KYM……危険予知ミーティングの略';

  function addKensaAnzen(wb) {
    var ws = portraitSheet(wb, sheetName('④-4 絶縁耐力試験 安全対策'));
    ws.columns = [{ width: 6 }, { width: 8 }, { width: 26 }, { width: 52 }];
    kensaHead(ws, '④-4', '絶縁耐力試験／安全対策及び注意事項', 2, 4);
    box(ws, 3, 1, 3, 4, KENSA_ANZEN, { size: 11, v: 'top' });
    ws.getRow(3).height = 560;
    return ws;
  }

  // [作業手順, 作業の要点, 注意事項及び安全対策]
  var KENSA_SHOCHI = [
    ['1) 作業前のミーティング\n　 KYMの実施\n　 役割分担確認\n　 試験内容の周知',
      '試験内容を把握させ危険を予測する。\n検査責任者が、試験責任者、安全管理責任者、試験指揮者、試験係員を点呼し、責任を明確にする。\n試験方法、時間を確認する。\n係員の配置を確認する。\n合図、連絡方法を確認する。',
      '試験範囲を全員に周知させる。\n試験対象となる電気工作物周辺の危険性を確認する。\n電圧印加範囲を関係者以外立入禁止とする。'],
    ['2) 試験前作業',
      '試験機器の配置、印加点を確認し準備する。\n短絡接地線を取り付ける。\n試験器操作場所と各電気室との連絡確認を行う。',
      '連絡時は電気室名と監視係の名前を確認し、誤りのないようにする。'],
    ['3) 試験回路のチェック',
      '印加区分図を使用して試験回路をチェックする。\n印加区分図を使用して処理した回路の復旧を確認する。',
      '印加区分図は試験指揮者が検査責任者とチェックする。'],
    ['4) 掲示物チェック', '必要な掲示物が揃っていることを確認する。', '絶縁耐力試験器設置箇所へ用意する。'],
    ['5) 備品チェック', '試験に必要な備品等をチェックする。また、試験終了後は、チェックリストにより機材、工具の確認を行う。',
      '試験前後に工具の品数が同じであることの確認をする。'],
    ['6) 人員チェック',
      '検電係・監視係が適正な配置がされていることを確認する。\n絶縁耐力試験終了後人員の点呼を行い全員無事であることの確認を行う。',
      '試験係員以外は立入禁止にする。\n点呼を必ず行う。'],
    ['7) 各所の異常の有無', '人員点呼後、異常の有無の報告を主任技術者に行う。', '終了後の異臭、変色等の有無の報告を行う。']
  ];

  function addKensaShochi(wb) {
    var ws = portraitSheet(wb, sheetName('④-5 処置確認チェックリスト'));
    ws.columns = [{ width: 6 }, { width: 16 }, { width: 31 }, { width: 31 }, { width: 8 }];
    var H = { size: 10, v: 'middle', h: 'center', fill: true };
    var T = { size: 10, v: 'top' };
    kensaHead(ws, '④-5', '絶縁耐力試験／処置確認チェックリスト', 2, 5);
    box(ws, 3, 1, 3, 2, '作業手順', H);
    box(ws, 3, 3, 3, 3, '作業の要点', H);
    box(ws, 3, 4, 3, 4, '注意事項及び安全対策', H);
    box(ws, 3, 5, 3, 5, 'チェック', H);
    ws.getRow(3).height = 20;
    KENSA_SHOCHI.forEach(function (t, i) {
      var r = 4 + i;
      box(ws, r, 1, r, 2, t[0], T);
      box(ws, r, 3, r, 3, t[1], T);
      box(ws, r, 4, r, 4, t[2], T);
      box(ws, r, 5, r, 5, '', { size: 10, v: 'middle', h: 'center' });
      ws.getRow(r).height = Math.max(36, kensaHeight(t[0], 22), kensaHeight(t[1], 31), kensaHeight(t[2], 31));
    });
    return ws;
  }

  // 保護継電器試験。[継電器名, [[試験名, 試験内容], ...]]。'【' で始まる試験名は要素の見出し
  var KENSA_KEIDENKI = [
    ['1. 過電流継電器', [
      ['(1)最小動作電流試験', '限時要素及び瞬時要素を整定タップに設定して測定する。'],
      ['(2)動作時間特性試験', '整定タップ、レバー10に設定して300%、700%の電流の動作時間を測定する。']]],
    ['2. 地絡過電流継電器', [
      ['(1)最小動作電流試験', '整定タップに設定して測定する。'],
      ['(2)動作時間特性試験', '整定タップの130%、400%の電流の動作時間を測定する。']]],
    ['3. 比率差動継電器', [
      ['(1)最小動作電流試験', '整定値において測定する。'],
      ['(2)動作時間特性試験', '整定値において0～300%まで電流を急変したときの動作時間を測定する。'],
      ['(3)比率特性試験', '1次又は2次の整定値の電流を一定にしたときの2次又は1次の動作電流を測定する。']]],
    ['4. 過電圧継電器、不足電圧継電器', [
      ['(1)最小（大）動作電圧試験', '整定タップにて測定する。'],
      ['(2)動作時間特性試験', '過電圧継電器は、整定タップの120%の電圧の動作時間を測定する。\n不足電圧継電器は、整定タップの70%の電圧の動作時間を測定する。']]],
    ['5. 地絡過電圧継電器', [
      ['(1)最小動作電圧試験', '整定値において測定する。'],
      ['(2)動作時間特性試験', '最小整定値、最大整定時間、整定電圧値の150%において動作時間を測定する。']]],
    ['6. 地絡方向継電器', [
      ['(1)最小動作電流試験', '整定タップに設定して、150%の電圧、動作位相の電流で測定する。'],
      ['(2)最小動作電圧試験', '整定タップに設定して、150%の電流、動作位相の電圧で測定する。'],
      ['(3)動作時間特性試験', '整定タップに設定して、150%の電圧として、130%、400%の電流の動作時間を測定する。'],
      ['(4)位相試験', '整定タップに設定して、150%電圧、1,000%の電流の動作位相角を測定する。']]],
    ['7. 漏電リレー（ELR）', [
      ['(1)タップ試験（最小動作電流試験）', '整定タップに設定して測定する。'],
      ['(2)動作時間特性試験', '整定タップに設定して、100%電流を入力した時の動作時間を測定する。']]],
    ['8. 絶縁監視装置', [
      ['(1)絶縁部試験（注意・警戒）', '管理値（注意15mA・警戒50mA）の10%増し電流値を試験整定値とし、試験模擬抵抗を接続した時の表示値を確認する。'],
      ['(2)漏電部試験', '整定タップに設定して測定する。']]],
    ['9. 3Eリレー', [
      ['【過負荷要素】', ''],
      ['(1)タップ試験（最小動作電流試験）', '整定タップに設定して測定する。'],
      ['(2)動作時間特性試験', '整定タップに設定して、600%電流を入力した時の動作時間を測定する。'],
      ['【欠相要素】', ''],
      ['(1)タップ試験（最小動作電流試験）', '整定タップに設定して測定する。'],
      ['(2)動作時間特性試験', '整定タップに設定して、100%電流を入力した時の動作時間を測定する（H要素）。'],
      ['【反相要素】', ''],
      ['(1)最小動作電圧試験', '定格電圧に設定して、徐々に電圧を低下させ測定する。'],
      ['(2)動作時間特性試験', '定格電圧に等しい反相電圧を入力した時の動作時間を測定する。']]]
  ];

  function addKensaKeidenki(wb) {
    var ws = portraitSheet(wb, sheetName('⑤-イ 保護継電器試験'));
    ws.pageSetup.fitToHeight = 0;             // 原本も3ページ。ページをまたいでも見出し行を印刷する
    ws.pageSetup.printTitlesRow = '4:4';
    ws.columns = [{ width: 6 }, { width: 8 }, { width: 38 }, { width: 40 }];
    var H = { size: 10, v: 'middle', h: 'center', fill: true };
    var M = { size: 10, v: 'middle' };
    var KIJUN = '製造者の管理値による。（試験成績表に記載）';
    var r = 5;
    function row(left, right, bold) {
      box(ws, r, 1, r, 3, left, { size: 10, v: 'middle', bold: bold });
      box(ws, r, 4, r, 4, right, M);
      ws.getRow(r).height = kensaHeight(left, 52);
      r++;
    }
    kensaHead(ws, '⑤-イ', '保護継電器試験', 2, 4);
    box(ws, 3, 1, 3, 1, '目的', H);
    box(ws, 3, 2, 3, 4, '継電器単体の特性に異常のないことを確認する。', M);
    ws.getRow(3).height = 24;
    box(ws, 4, 1, 4, 3, '検査内容', H);
    box(ws, 4, 4, 4, 4, '判定基準', H);
    ws.getRow(4).height = 20;
    row('製造者の示す動作特性試験点で試験を行い、継電器単体の良否を確認する。', '', false);
    KENSA_KEIDENKI.forEach(function (k) {
      row(k[0], '', true);
      k[1].forEach(function (t) {
        if (t[0].charAt(0) === '【') { row('　' + t[0], '', false); } else { row(t[0] + '\n　' + t[1], KIJUN, false); }
      });
    });
    box(ws, r, 1, r, 1, '注意\n事項', H);
    box(ws, r, 2, r, 4, '1. 引き出し形の継電器の場合は引き出す際の落下に注意する。\n2. テストプラグの短絡バーの使用状態を確認する。\n' +
      '3. 試験電圧を印加したままテストプラグを脱着しないこと。', M);
    ws.getRow(r).height = 54;
    box(ws, r + 1, 1, r + 1, 1, '備考', H);
    box(ws, r + 1, 2, r + 1, 4, '試験記録は「保護継電器試験成績表」に結果を記入する。', M);
    ws.getRow(r + 1).height = 30;
    return ws;
  }

  // ⑤-ロ・⑤-ハ（⑤-イ のあとに出力する）
  var KENSA_GO = [
    { no: '⑤-ロ', title: 'インターロック試験',
      mokuteki: '各機器、装置の動作に異常のないことを確認する。',
      naiyo: 'インターロック条件を成立させて機器の「入」「切」が可能なことを、不成立にして機器の「入」「切」ができないことを確認する。',
      kijun: '製造者の社内規格による試験方法により製作図にあるインターロック図に従って動作をすること、又は動作しないこと。',
      chui: '1. 機器操作時は、誤操作しないよう注意する。\n2. 連絡方法を確認する。',
      biko: '試験記録は「インターロック試験成績表」に結果を記入する。' },
    { no: '⑤-ハ', title: '保護連動試験',
      mokuteki: '各機器・装置の動作に異常のないことを確認する。',
      naiyo: '(1)保護連動、故障表示試験\n　保護継電器を模擬動作させ、遮断器のトリップ及び故障表示・故障警報を現場盤及び監視盤にて確認する。\n\n' +
        '(2)中央監視動作試験\n　現場盤より模擬の故障等を発生させ、故障表示・故障警報を確認する。中央監視装置より機器操作を行い、表示を確認する。',
      kijun: '製造者の社内規格による試験方法により、製作図にある保護連動表に従って動作すること。',
      chui: '1. 機器操作時は、誤操作しないよう注意する。\n2. 連絡方法を確認する。',
      biko: '試験記録は「保護連動試験成績表」に結果を記入する。' }
  ];

  function addKensa(wb) {
    KENSA.forEach(function (k) {
      if (k.calc) { addKensaBshu(wb, k); } else { addKensaSheet(wb, k); }
    });
    addKensaTejun(wb);
    addKensaFlow(wb);
    addKensaAnzen(wb);
    addKensaShochi(wb);
    addKensaKeidenki(wb);
    KENSA_GO.forEach(function (k) { addKensaSheet(wb, k); });
  }

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
    if (opt.hannyu === true) { var kq = Number(state.koshu) || 0;
      addHannyu(wb, info, kq);
      addHannyuTaisei(wb, info, kq);
      // 太陽光発電装置（作成例 (10)）には搬入経路・ワイヤロープ・アンカーの項が無いため出力しない（利用者の指示）
      if (kq !== 9) {
        addHannyuZu(wb, info, kq);
        addWire(wb, info, kq);
      }
      addSuetsuke(wb, info, kq);
      if (kq !== 9) { addAnchor(wb, info); }
    }
    if (opt.jishu === true) { addKensa(wb); }
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
