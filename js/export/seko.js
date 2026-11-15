/* 施工計画書 Excel(xlsx) 出力：電気設備工事施工チェックシート（公共建築協会 令和7年版）・申請手続一覧表
 * データは data/sekocheck.js（window.SEKOCHECK_DATA。tools/build_sekocheck.py で docs/sekocheck/*.md から自動生成）
 * 分割したファイルは index.html の順（core → … → main）に読み込む。部品は window.SKUKIKKSExport._ で共有する
 */
(function (X) {
  'use strict';

  var FONT = X.FONT, box = X.box, lineCount = X.lineCount, portraitSheet = X.portraitSheet, sheetName = X.sheetName,
    textLine = X.textLine, pageSetup = X.pageSetup, wareki = X.wareki;

  var ROW_MAX = 400;     // 1行の高さの上限(pt)。Excel の行の高さは 409pt までのため、これを超える項目だけ続きの行に分ける
  var MERGE_MAX = 420;   // 申請手続一覧表の工事区分の縦結合を区切る高さ(pt)
  var NO = { ippan: '1', haikan: '2-1', haisen: '2-2', kiki: '2-3', setchi: '2-4', rai: '2-5', kakuu: '2-6',
    chichu: '2-7', juhen: '3', kanshi: '4', iryo: '5' };

  function data() { return window.SEKOCHECK_DATA || { chapters: [] }; }

  /* ---------- 表紙 ---------- */

  function addSekoCover(wb, info) {
    var ws = portraitSheet(wb, '施工チェックシート 表紙');
    ws.columns = [{ width: 4 }, { width: 18 }, { width: 8 }, { width: 46 }, { width: 4 }];
    var i;
    for (i = 1; i <= 4; i++) { ws.getRow(i).height = 30; }
    textLine(ws, 5, 5, '電気設備工事施工チェックシート', 26, true, 60);
    ws.getRow(6).height = 60;
    var L = { size: 12, v: 'middle', h: 'distributed', border: false };
    var V = { size: 12, v: 'middle', border: false };
    var UL = { style: 'thin', color: { argb: 'FF000000' } };
    var r = 7;
    function line(label, sub, value) {
      box(ws, r, 2, r, 2, label, L);
      box(ws, r, 3, r, 3, sub || '', { size: 12, v: 'middle', h: 'center', border: false });
      box(ws, r, 4, r, 4, value || '', V);
      ws.getCell(r, 4).border = { bottom: UL };
      if (!sub) { ws.getCell(r, 3).border = { bottom: UL }; }
      ws.getRow(r).height = 34;
      r++;
    }
    var blankDate = '令和　　年　　月　　日';
    line('工事名称', '', info.kojiName);
    line('工事場所', '', info.basho);
    line('建築物名称', '', '');
    line('工　　期', '自', info.kokiFrom ? wareki(info.kokiFrom) : blankDate);
    line('', '至', info.kokiTo ? wareki(info.kokiTo) : blankDate);
    line('', '変更', blankDate);
    ws.getRow(r).height = 30; r++;
    line('受注者', '', info.juchu);
    line('住所・電話', '', '');
    line('', '', '');
    ws.getRow(r).height = 60; r++;
    textLine(ws, r, 5, '出典：一般社団法人 公共建築協会「電気設備工事施工チェックシート 令和7年版」', 9, false, 18);
    return ws;
  }

  /* ---------- チェックシートの使用にあたって ---------- */

  var TSUKAI = [
    '1.　工程ごとのチェックポイントを、準備段階と施工段階とに区分し、準備段階ではあらかじめ打合せする事項を示し、施工段階では直接施工についてチェックする事項を記載しています。\n　なお、特記仕様書等で示された事項は、適宜チェックポイントを追加して下さい。',
    '2.　チェックポイントの項に(検)又は(立)と表示してある事項は、「公共建築工事標準仕様書（電気設備工事編）令和7年版」（以下「標準仕様書」という。）による監督職員の検査又は立会いが定められていることを示しています。ただし、一般共通事項による検査及び立会いは、各施工内容に応じて適切に行うものとし、特に表示をしていません。',
    '3.　チェックシートの項（1-1.1.5）などの表示は、該当する「標準仕様書」の編・章・節・項の番号を示しています。',
    '4.　チェックシートの使い方\n　1)　本シートは、原則として建築物1棟ごとに使用します。\n　2)　本シートの具体的な使用方法\n　　⑴　準備段階においては、チェックポイントに従い、契約条件、工程、施工計画・施工図等について、打合せを行い、その内容等を適宜記録欄に記入します。\n　　⑵　施工段階においては、チェックポイントに従い、施工内容等の確認を行い、その内容を適宜記録欄に記入します。\n　　⑶　記録欄には、チェックの日付、打合せ事項、指示・承諾・確認事項等を記入します。\n　　⑷　繰り返される工事などで記録欄に記入しきれない場合は、Excel で行を追加して記入してください。',
    '5.　本シートの使用にあたっては、下記の参考図書があります。\n　・「営繕工事写真撮影要領による 工事写真撮影ガイドブック　電気設備工事編　令和5年版」\n　　監修 国土交通省大臣官房官庁営繕部／編集・発行（一社）公共建築協会\n　・「電気設備工事監理指針（令和7年版）」\n　　監修 国土交通省大臣官房官庁営繕部／編集（一社）公共建築協会／発行（一社）建設電気技術協会'
  ];

  function addSekoTsukai(wb) {
    var ws = portraitSheet(wb, '使用にあたって');
    ws.columns = [{ width: 88 }];
    textLine(ws, 1, 1, 'チェックシートの使用にあたって', 16, true, 34);
    ws.getRow(2).height = 10;
    TSUKAI.forEach(function (t, i) {
      var c = ws.getCell(3 + i, 1);
      c.value = t;
      c.font = { name: FONT, size: 10.5 };
      c.alignment = { wrapText: true, vertical: 'top', horizontal: 'left' };
      ws.getRow(3 + i).height = lineCount(t, 88, 10.5) * 15 + 10;
    });
    return ws;
  }

  /* ---------- チェックシート本体（章・段階ごとに1シート） ---------- */

  // 項番号を表示用の行にする（1行に1つ。途中で折り返さないため）。小項目（①…）付きは番号の最初の行にだけ記号を付ける
  function refLines(refs) {
    var out = [];
    var last = '';
    refs.forEach(function (r) {
      var m = /^([\u2460-\u2473\u3251-\u3255])(.+)$/.exec(r);
      if (m) {
        out.push((m[1] === last ? '　 ' : m[1] + ' ') + m[2]);
        last = m[1];
      } else {
        out.push(r);
        last = '';
      }
    });
    return out;
  }

  // 行の高さ(pt)の見積もり（チェックポイント・項番号の列のうち高い方）
  function rowHeight(t, refs) {
    return Math.ceil(Math.max(lineCount(t, 44, 10) * 13.5, lineCount(refs.join('\n'), 16, 9) * 12.5) + 8);
  }

  // 1項目を1行にする。Excel の行の高さの上限を超える項目（配管工事の「配管図」など）だけ、小項目の区切りで
  // 続きの行（「（続き）」）に分ける。戻り値 [{t, ref:[...]}]
  function splitItem(it) {
    var all = refLines(it.ref);
    if (rowHeight(it.t, all) <= ROW_MAX) { return [{ t: it.t, ref: all }]; }
    var lines = it.t.split('\n');
    var head = lines[0];
    var rows = [];
    var part = [];
    function refsFor(subs, first) {
      var marks = subs.map(function (x) { return x.charAt(1); });
      return refLines(it.ref.filter(function (r) {
        var m = r.charAt(0);
        return marks.indexOf(m) >= 0 || (first && !/[\u2460-\u2473\u3251-\u3255]/.test(m));
      }));
    }
    function text(subs, first) { return (first ? head : head + '（続き）') + '\n' + subs.join('\n'); }
    lines.slice(1).forEach(function (ln) {
      var first = rows.length === 0;
      var next = part.concat([ln]);
      if (part.length && rowHeight(text(next, first), refsFor(next, first)) > ROW_MAX) {
        rows.push({ t: text(part, first), ref: refsFor(part, first) });
        part = [ln];
      } else {
        part = next;
      }
    });
    if (part.length) { rows.push({ t: text(part, rows.length === 0), ref: refsFor(part, rows.length === 0) }); }
    return rows;
  }

  function addSekoSheet(wb, info, ch, sh) {
    var stage = sh.stage ? '　' + sh.stage : '';
    var name = sheetName((NO[ch.id] || '') + ' ' + ch.name + (sh.stage ? ' ' + sh.stage.replace('段階', '') : ''));
    var ws = wb.addWorksheet(name, {
      pageSetup: pageSetup('portrait', '3:3'),
      views: [{ state: 'frozen', ySplit: 3 }],
      headerFooter: { oddFooter: '&C&P / &N' }
    });
    var WA = 13, WB = 44, WC = 16, WD = 30;
    ws.columns = [{ width: WA }, { width: WB }, { width: WC }, { width: WD }];
    textLine(ws, 1, 4, ch.title + stage, 12, true, 24, 'left');
    textLine(ws, 2, 4, '工事名　' + (info.kojiName || ''), 9, false, 18, 'left');
    var H = { size: 10, v: 'middle', h: 'center', fill: true };
    box(ws, 3, 1, 3, 3, 'チ ェ ッ ク ポ イ ン ト', H);
    box(ws, 3, 4, 3, 4, '記　　録', H);
    ws.getRow(3).height = 20;

    var r = 4;
    sh.kubun.forEach(function (k) {
      var label = (k.mark ? '(' + k.mark + ')\n' : '') + k.name;
      var rows = [];
      k.items.forEach(function (it) { rows = rows.concat(splitItem(it)); });
      var start = r;
      rows.forEach(function (row) {
        var refText = row.ref.join('\n');
        var n = Math.max(lineCount(row.t, WB, 10) * 13.5, lineCount(refText, WC, 9) * 12.5,
          lineCount(label, WA, 10) * 13.5);
        var h = Math.min(409, Math.max(26, Math.ceil(n + 8)));
        box(ws, r, 2, r, 2, row.t, { size: 10, v: 'top' });
        box(ws, r, 3, r, 3, refText, { size: 9, v: 'top' });
        box(ws, r, 4, r, 4, '', { size: 10, v: 'top' });
        ws.getRow(r).height = h;
        r++;
      });
      box(ws, start, 1, r - 1, 1, label, { size: 10, v: 'top' });
    });
    return ws;
  }

  // 施工チェックシート一式（表紙・使用にあたって・選んだ章）。chaps は { 章id: true } （未指定はすべて）
  function addSeko(wb, info, chaps) {
    addSekoCover(wb, info);
    addSekoTsukai(wb);
    data().chapters.forEach(function (ch) {
      if (chaps && !chaps[ch.id]) { return; }
      ch.sheets.forEach(function (sh) { addSekoSheet(wb, info, ch, sh); });
    });
  }

  /* ---------- 参考資料 主な官公署その他への申請手続一覧表 ---------- */

  // 同じ値が続く行を縦に結合する（高さ MERGE_MAX ごとに区切る）
  function mergeSame(ws, r0, rows, col, heights, same, opt) {
    var s = 0;
    var acc = heights[0];
    var i;
    for (i = 1; i <= rows.length; i++) {
      var cont = i < rows.length && same(rows[i], rows[s]) && acc + heights[i] <= MERGE_MAX;
      if (cont) { acc += heights[i]; continue; }
      box(ws, r0 + s, col, r0 + i - 1, col, rows[s][col - 1], opt);
      s = i;
      acc = i < rows.length ? heights[i] : 0;
    }
  }

  function addShinsei(wb, info) {
    var d = data().shinsei;
    if (!d) { return null; }
    var W = [7, 9, 24, 13, 15, 13, 30, 19];
    var ws = wb.addWorksheet('申請手続一覧表', {
      pageSetup: pageSetup('landscape', '3:3'),
      views: [{ state: 'frozen', ySplit: 3 }],
      headerFooter: { oddFooter: '&C&P / &N' }
    });
    ws.columns = W.map(function (w) { return { width: w }; });
    textLine(ws, 1, 8, '参考資料　' + d.title, 14, true, 26, 'left');
    textLine(ws, 2, 8, '工事名　' + (info.kojiName || ''), 9, false, 18, 'left');
    var H = { size: 9, v: 'middle', h: 'center', fill: true };
    box(ws, 3, 1, 3, 2, '工事区分', H);
    d.heads.slice(2).forEach(function (t, i) { box(ws, 3, 3 + i, 3, 3 + i, t, H); });
    ws.getRow(3).height = 20;
    var r0 = 4;
    var heights = d.rows.map(function (row) {
      var n = 1;
      row.forEach(function (v, i) { if (i >= 2) { n = Math.max(n, lineCount(v, W[i], 9)); } });
      return Math.max(24, n * 12.5 + 6);
    });
    d.rows.forEach(function (row, i) {
      var r = r0 + i;
      row.forEach(function (v, c) {
        if (c >= 2) { box(ws, r, c + 1, r, c + 1, v, { size: 9, v: 'top' }); }
      });
      ws.getRow(r).height = heights[i];
    });
    var V = { size: 9, v: 'middle', h: 'center', vertical: true };
    mergeSame(ws, r0, d.rows, 1, heights, function (a, b) { return a[0] === b[0]; }, V);
    mergeSame(ws, r0, d.rows, 2, heights, function (a, b) { return a[0] === b[0] && a[1] === b[1]; },
      { size: 9, v: 'middle', h: 'center' });
    var r = r0 + d.rows.length;
    ws.getRow(r).height = 8;
    d.notes.forEach(function (t) {
      r++;
      textLine(ws, r, 8, '注' + t, 9, false, 16, 'left');
    });
    return ws;
  }

  function addDenjisoku(wb, info) {
    var d = data().shinsei;
    if (!d || !d.beppyo) { return null; }
    var b = d.beppyo;
    var W = [16, 26, 76];
    var ws = wb.addWorksheet('電事則別表第二', {
      pageSetup: pageSetup('landscape', '3:3'),
      headerFooter: { oddFooter: '&C&P / &N' }
    });
    ws.columns = W.map(function (w) { return { width: w }; });
    textLine(ws, 1, 3, b.title, 13, true, 24, 'left');
    textLine(ws, 2, 3, '（申請手続一覧表 注3。工事計画届出の適用範囲）', 9, false, 18, 'left');
    var H = { size: 9, v: 'middle', h: 'center', fill: true };
    box(ws, 3, 1, 3, 2, '工事の種類', H);
    box(ws, 3, 3, 3, 3, '事前届出を要するもの', H);
    ws.getRow(3).height = 20;
    var r0 = 4;
    b.rows.forEach(function (row, i) {
      var r = r0 + i;
      box(ws, r, 2, r, 2, row[0], { size: 9, v: 'top' });
      box(ws, r, 3, r, 3, row[1], { size: 9, v: 'top' });
      ws.getRow(r).height = Math.max(22, Math.max(lineCount(row[0], W[1], 9), lineCount(row[1], W[2], 9)) * 12.5 + 6);
    });
    box(ws, r0, 1, r0 + b.rows.length - 1, 1, b.kind, { size: 9, v: 'top' });
    return ws;
  }

  // ほかのファイルから使う部品
  X.addSeko = addSeko;
  X.addShinsei = addShinsei;
  X.addDenjisoku = addDenjisoku;
  X.sekoChapters = function () {
    return data().chapters.map(function (ch) { return { id: ch.id, no: NO[ch.id] || '', name: ch.name }; });
  };
})(window.SKUKIKKSExport._);
