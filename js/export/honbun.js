/* 施工計画書 Excel(xlsx) 出力：本文シート（data/honbun.js は本文を出力するときだけ読み込む）
 * 分割したファイルは index.html の順（core → … → main）に読み込む。部品は window.SKUKIKKSExport._ で共有する
 */
(function (X) {
  'use strict';

  var FONT = X.FONT, KOSHU = X.KOSHU, box = X.box, lineCount = X.lineCount, portraitSheet = X.portraitSheet,
    textLine = X.textLine, vlen = X.vlen;

  /* ---------- 本文（作成例の 1.〜7. の文章。data/honbun.js） ---------- */

  // 総合施工計画書（sogo: true）か、工種の本文を出力する。工事名は入力値に置き換える
  function addHonbun(wb, info, koshu, sogo) {
    var all = window.HONBUN_DATA || {};
    var chaps = sogo ? all.sogo : (all.koshu || {})[koshu];
    if (!chaps || !chaps.length) { return null; }
    var NCOL = 12;
    var ws = portraitSheet(wb, '本文');
    ws.pageSetup.fitToHeight = 0;
    var cols = [];
    var i;
    for (i = 0; i < NCOL; i++) { cols.push({ width: 7.5 }); }
    ws.columns = cols;
    var koji = info.kojiName || '';
    function rep(t) { return koji ? String(t).replace(/○○○○○○新築電気設備工事/g, koji) : String(t); }
    var title = sogo ? '総合施工計画書' : (koshu < KOSHU.length ? KOSHU[koshu] : (info.docTitle || ''));
    var r = 1;
    textLine(ws, r, NCOL, title, 14, true, 28); r++;
    textLine(ws, r, NCOL, '工事名　' + koji, 10, false, 20, 'left'); r++;
    function para(text, lv, size, bold, color) {
      ws.mergeCells(r, 1, r, NCOL);
      var c = ws.getCell(r, 1);
      c.value = text;
      c.font = { name: FONT, size: size, bold: !!bold, color: color ? { argb: color } : undefined };
      c.alignment = { vertical: 'top', wrapText: true, indent: Math.min(lv * 2, 10) };
      var w = NCOL * 7.5 - lv * 2.2;
      ws.getRow(r).height = Math.max(size + 8, Math.ceil(lineCount(text, w, size) * size * 1.4 + 5));
      r++;
    }
    function table(rows, lv) {
      var n = Math.min(rows[0].length, NCOL);
      var c0 = 1 + Math.min(lv, 2);
      var span = NCOL - c0 + 1;
      // 列の幅は文字数の多さで配分する（各列 1 以上）
      var len = [];
      var k, q;
      for (k = 0; k < n; k++) {
        len[k] = 2;
        rows.forEach(function (row) { len[k] = Math.max(len[k], Math.min(40, vlen(String(row[k] || '')))); });
      }
      var tot = len.reduce(function (a, b) { return a + b; }, 0);
      var w = [];
      var used = 0;
      for (k = 0; k < n; k++) {
        w[k] = k === n - 1 ? span - used : Math.max(1, Math.round(len[k] / tot * span));
        if (k < n - 1 && used + w[k] > span - (n - 1 - k)) { w[k] = span - (n - 1 - k) - used; }
        used += w[k];
      }
      rows.forEach(function (row, ri) {
        var c = c0;
        var lines = 1;
        for (q = 0; q < n; q++) {
          var t = rep(row[q] || '');
          box(ws, r, c, r, c + w[q] - 1, t, { size: 9, v: 'middle', h: ri === 0 ? 'center' : 'left', fill: ri === 0 });
          lines = Math.max(lines, lineCount(t, w[q] * 7.5, 9));
          c += w[q];
        }
        ws.getRow(r).height = Math.max(18, lines * 12.5 + 6);
        r++;
      });
    }
    chaps.forEach(function (ch) {
      ws.getRow(r).height = 8; r++;
      para(ch.no + '. ' + ch.title, 0, 12, true);
      ch.lines.forEach(function (l) {
        var lv = l.lv + 1;
        if (l.t === 'tb') { table(l.rows, lv); return; }
        if (l.t === 'fig') { para(l.text, lv, 9, false, 'FF808080'); return; }
        var t = rep(l.text);
        var head = /^\d+\.\d+ /.test(t);
        para((head || l.t === 'cap' || /^[①-⑳⑴-⑽]/.test(t) || /^[（(]/.test(t) || /^※/.test(t)) ? t : '・' + t,
          lv, head ? 11 : 10, head || l.t === 'cap');
      });
    });
    return ws;
  }

  // 本文データ（約170KB）は本文を出力するときだけ読み込む
  var honbunLoading = null;
  function loadHonbun() {
    if (window.HONBUN_DATA) { return Promise.resolve(window.HONBUN_DATA); }
    if (!honbunLoading) {
      honbunLoading = new Promise(function (resolve, reject) {
        var s = document.createElement('script');
        s.src = 'data/honbun.js';
        s.onload = function () { resolve(window.HONBUN_DATA); };
        s.onerror = function () { honbunLoading = null; reject(new Error('本文データを読み込めませんでした')); };
        document.head.appendChild(s);
      });
    }
    return honbunLoading;
  }

  // 本文は「工種別」のときだけ工種の施工計画書、それ以外は総合施工計画書の本文にする
  function honbunIsSogo(state) { return state.docType !== 'koshu'; }

  // ほかのファイルから使う部品
  X.addHonbun = addHonbun;
  X.honbunLoading = honbunLoading;
  X.loadHonbun = loadHonbun;
  X.honbunIsSogo = honbunIsSogo;
})(window.SKUKIKKSExport._);
