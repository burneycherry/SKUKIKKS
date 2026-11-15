/* 施工計画書 Excel(xlsx) 出力：品質管理チェックシート・工事写真の撮影箇所
 * 分割したファイルは index.html の順（core → … → main）に読み込む。部品は window.SekokanExport._ で共有する
 */
(function (X) {
  'use strict';

  var KOSHU = X.KOSHU, box = X.box, lineCount = X.lineCount, portraitSheet = X.portraitSheet, sheetName = X.sheetName,
    textLine = X.textLine, SHIUN = X.SHIUN;

  /* ---------- 品質管理チェックシート（工種別施工計画書の別添1） ---------- */

  function addCheckSheet(wb, info, koshuName, no, sheet) {
    var ws = portraitSheet(wb, sheetName('チェック' + no + ' ' + (sheet.name || sheet.item || '様式')));
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
    if (koshu === 14) {                 // 総合試運転計画書は、作成例 (15) の試験ごとのチェックリストを使う
      sheets = [];
      SHIUN.forEach(function (s) {
        s.checks.forEach(function (ck) {
          sheets.push({ title: ck.cap, name: ck.cap.replace(/^【|】.*$/g, ''), item: ck.item.replace(/\n/g, ''), rows: ck.rows });
        });
      });
    }
    if (!sheets || !sheets.length) {
      var blank = [];
      var i;
      for (i = 0; i < 12; i++) { blank.push(['', '', '']); }
      sheets = [{ title: '', item: '', rows: blank }];
    }
    sheets.forEach(function (sh, i) { addCheckSheet(wb, info, name, i + 1, sh); });
  }

  /* ---------- 工事写真の撮影箇所（工種別の別添2。「営繕工事写真撮影要領」撮影対象表（電気設備工事編）の抜粋） ---------- */

  function addSatsuei(wb, info, koshu) {
    var all = window.SATSUEI_DATA || {};
    if (!(koshu >= 0 && koshu <= KOSHU.length)) { koshu = 0; }
    var name = koshu < KOSHU.length ? KOSHU[koshu].replace(/施工計画書$|計画書$/, '') : (info.docTitle || '');
    var rows = all[koshu];
    var nuki = !!(rows && rows.length);
    if (!nuki) { rows = all.all || []; }
    var ws = portraitSheet(wb, '工事写真の撮影箇所');
    ws.pageSetup.fitToHeight = 0;
    ws.pageSetup.printTitlesRow = '4:4';
    ws.columns = [{ width: 15 }, { width: 17 }, { width: 50 }, { width: 10 }];
    var H = { size: 10, v: 'middle', h: 'center', fill: true };
    var C = { size: 10, v: 'middle' };
    textLine(ws, 1, 4, '工事写真の撮影箇所', 16, true, 30);
    textLine(ws, 2, 4, '工事名　' + (info.kojiName || '') + '　／　' + name, 10, false, 20, 'left');
    textLine(ws, 3, 4, '（「営繕工事写真撮影要領」令和5年版' + (nuki ? 'の抜粋' : '') + '）撮影対象表（電気設備工事編）', 10, false, 20, 'left');
    ['工事種目又は分類', '撮影項目', '撮影対象', '撮影時期'].forEach(function (t, i) { box(ws, 4, 1 + i, 4, 1 + i, t, H); });
    ws.getRow(4).height = 20;
    var r0 = 5;
    rows.forEach(function (row, i) {
      var r = r0 + i;
      box(ws, r, 1, r, 1, row[0], { size: 10, v: 'top' });
      box(ws, r, 2, r, 2, row[1], { size: 10, v: 'top' });
      box(ws, r, 3, r, 3, row[2], C);
      box(ws, r, 4, r, 4, row[3], { size: 10, v: 'middle', h: 'center' });
      ws.getRow(r).height = Math.max(22, lineCount(row[2], 50, 10) * 13.5 + 8);
    });
    // 同じ工事種目・撮影項目が続く行は縦に結合する（原本と同じ）。全行を出すときはページをまたぐため結合しない
    if (nuki) {
      [0, 1].forEach(function (col) {
        var s = 0;
        var i;
        for (i = 1; i <= rows.length; i++) {
          var same = i < rows.length && rows[i][col] === rows[s][col] && (col === 0 || rows[i][0] === rows[s][0]);
          if (!same) {
            if (i - 1 > s) { ws.mergeCells(r0 + s, col + 1, r0 + i - 1, col + 1); }
            s = i;
          }
        }
      });
    }
    var e = r0 + rows.length;
    ws.getRow(e).height = 8;
    textLine(ws, e + 1, 4, '※受注者が作成した資料を添付する場合は、営繕工事写真撮影要領を満たしていることを記載する。', 9, false, 18, 'left');
    return ws;
  }

  // ほかのファイルから使う部品
  X.addCheckSheet = addCheckSheet;
  X.addChecks = addChecks;
  X.addSatsuei = addSatsuei;
})(window.SekokanExport._);
