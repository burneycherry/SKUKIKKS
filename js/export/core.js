/* 施工計画書 Excel(xlsx) 出力：共通の部品（文字列・書式・シートの作成・罫線・一覧表、選択内容の収集、和暦、工種の一覧 KOSHU）
 * 分割したファイルは index.html の順（core → … → main）に読み込む。部品は window.SKUKIKKSExport._ で共有する
 */
window.SKUKIKKSExport = { _: {} };

(function (X) {
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

  // シート名に使えない文字を除き、31文字以内にする
  function sheetName(s) {
    return String(s).replace(/[\[\]:*?\/\\：]/g, ' ').substring(0, 31);
  }

  /* ---------- 一覧表（汎用） ---------- */

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

  // ほかのファイルから使う部品
  X.FONT = FONT;
  X.FS = FS;
  X.LH = LH;
  X.CHUNK = CHUNK;
  X.THIN = THIN;
  X.BORDER = BORDER;
  X.HEAD_FILL = HEAD_FILL;
  X.W = W;
  X.vlen = vlen;
  X.lineCount = lineCount;
  X.heightFor = heightFor;
  X.wareki = wareki;
  X.parseDocs = parseDocs;
  X.collect = collect;
  X.style = style;
  X.pageSetup = pageSetup;
  X.portraitSheet = portraitSheet;
  X.totalWidth = totalWidth;
  X.textLine = textLine;
  X.box = box;
  X.orMaru = orMaru;
  X.vline = vline;
  X.hline = hline;
  X.landscapeSheet = landscapeSheet;
  X.addBlank = addBlank;
  X.KOSHU = KOSHU;
  X.sheetName = sheetName;
  X.addTable = addTable;
})(window.SKUKIKKSExport._);
