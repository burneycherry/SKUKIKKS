/* 施工計画書 Excel(xlsx) 出力
 * 依存: vendor/exceljs.min.js (window.ExcelJS)
 * 公開: window.SekokanExport = { build, collect, parseDocs, wareki }
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

  /* ---------- 表紙 ---------- */

  function addCover(wb, info, chapters) {
    var ws = wb.addWorksheet('表紙', {
      pageSetup: {
        paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 1,
        horizontalCentered: true,
        margins: { left: 0.7, right: 0.7, top: 0.8, bottom: 0.8, header: 0.3, footer: 0.3 }
      }
    });
    ws.columns = [{ width: 4 }, { width: 20 }, { width: 58 }, { width: 4 }];

    function line(row, text, size, bold, height) {
      ws.mergeCells(row, 1, row, 4);
      var c = ws.getCell(row, 1);
      c.value = text;
      c.font = { name: FONT, size: size, bold: bold };
      c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      ws.getRow(row).height = height;
    }

    ws.getRow(1).height = 40;
    ws.getRow(2).height = 40;
    line(3, info.docTitle || '施工計画書', 32, true, 60);
    line(4, '（電気設備工事）', 16, false, 32);
    ws.getRow(5).height = 40;

    var rows = [
      ['工事名', info.kojiName],
      ['工事場所', info.basho],
      ['工期', (info.kokiFrom || info.kokiTo)
        ? (wareki(info.kokiFrom) + ' ～ ' + wareki(info.kokiTo)) : ''],
      ['発注者', info.hacchu],
      ['受注者', info.juchu],
      ['現場代理人', info.dairinin],
      ['監理(主任)技術者', info.gijutsu],
      ['作成者', info.sakuseisha],
      ['作成年月日', wareki(info.sakuseiBi)]
    ];
    var r = 6;
    rows.forEach(function (kv) {
      var a = ws.getCell(r, 2);
      var b = ws.getCell(r, 3);
      a.value = kv[0];
      b.value = kv[1] || '';
      style(a, { size: 12, v: 'middle', h: 'center', fill: true });
      style(b, { size: 12, v: 'middle' });
      ws.getRow(r).height = Math.max(30, lineCount(kv[1] || '', 58, 12) * 17 + 8);
      r++;
    });

    ws.getRow(r).height = 30;
    r++;
    var h = ws.getCell(r, 2);
    h.value = '対象工事種目';
    ws.mergeCells(r, 2, r, 3);
    style(h, { size: 11, v: 'middle', h: 'center', fill: true });
    style(ws.getCell(r, 3), { size: 11, fill: true });
    ws.getRow(r).height = 24;
    r++;
    var names = chapters.map(function (ch) { return ch.title.replace(/　/g, ' '); });
    var body = ws.getCell(r, 2);
    ws.mergeCells(r, 2, r, 3);
    body.value = names.join('\n');
    style(body, { size: 11, v: 'middle' });
    style(ws.getCell(r, 3), { size: 11 });
    ws.getRow(r).height = Math.max(30, names.length * 16 + 10);
    r += 2;

    line(r, '本書の管理表は、国土交通省「電気設備工事 施工管理 管理表（令和8年版）」から対象の工事種目を抜粋したものです。',
      9, false, 30);
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
    if (!chapters.length) { return Promise.reject(new Error('工事種目が選択されていません')); }

    var wb = new ExcelJS.Workbook();
    wb.creator = info.juchu || 'sekokan';
    wb.created = new Date();

    if (opt.cover !== false) { addCover(wb, info, chapters); }
    if (opt.list !== false) { addItemList(wb, info, chapters); }
    chapters.forEach(function (ch) { addChapter(wb, info, ch, opt.page !== false); });
    if (opt.docs !== false) { addDocList(wb, info, chapters); }
    if (opt.photos !== false) { addPhotoList(wb, info, chapters); }

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
    wareki: wareki
  };
})();
