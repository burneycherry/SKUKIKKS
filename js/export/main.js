/* 施工計画書 Excel(xlsx) 出力：ブック全体の組み立てと公開 API（window.SKUKIKKSExport）
 * 分割したファイルは index.html の順（core → … → main）に読み込む。部品は window.SKUKIKKSExport._ で共有する
 */
(function (X) {
  'use strict';

  var KOSHU = X.KOSHU, addBlank = X.addBlank, collect = X.collect, parseDocs = X.parseDocs, wareki = X.wareki,
    addAnzenKeikaku = X.addAnzenKeikaku, addAnzenSoshiki = X.addAnzenSoshiki, addCover = X.addCover,
    addGaiyo = X.addGaiyo, addHinshitsu = X.addHinshitsu, addHoan = X.addHoan, addKoshu = X.addKoshu,
    addRenraku = X.addRenraku, addSoshiki = X.addSoshiki, addChapter = X.addChapter, addDocList = X.addDocList,
    addItemList = X.addItemList, addPhotoList = X.addPhotoList, addHonbun = X.addHonbun,
    honbunIsSogo = X.honbunIsSogo, loadHonbun = X.loadHonbun, addAnchor = X.addAnchor, addHannyu = X.addHannyu,
    addHannyuTaisei = X.addHannyuTaisei, addHannyuZu = X.addHannyuZu, addSuetsuke = X.addSuetsuke,
    addWire = X.addWire, addKensa = X.addKensa, addHokoku = X.addHokoku, addShiunten = X.addShiunten,
    addChecks = X.addChecks, addSatsuei = X.addSatsuei, addSeko = X.addSeko, addShinsei = X.addShinsei,
    addDenjisoku = X.addDenjisoku;

  /* ---------- ブック全体 ---------- */

  function build(data, state) {
    var need = state && state.opt && state.opt.honbun === true;
    return (need ? loadHonbun() : Promise.resolve()).then(function () { return buildNow(data, state); });
  }

  function buildNow(data, state) {
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
    wb.creator = info.juchu || 'SKUKIKKS';
    wb.created = new Date();

    if (opt.cover !== false) { addCover(wb, info, state.docType === 'koshu'); }
    if (opt.honbun === true) { addHonbun(wb, info, Number(state.koshu) || 0, honbunIsSogo(state)); }
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
    if (opt.shiunten === true) { addShiunten(wb, info); }
    if (opt.hokoku === true) { addHokoku(wb, info); }
    if (opt.check === true) { addChecks(wb, info, Number(state.koshu) || 0); }
    if (opt.satsuei === true) { addSatsuei(wb, info, Number(state.koshu) || 0); }
    if (opt.seko === true) { addSeko(wb, info, state.sekoChap || null); }
    if (opt.shinsei === true) { addShinsei(wb, info); addDenjisoku(wb, info); }
    if (!wb.worksheets.length) { return Promise.reject(new Error('出力するシートがありません。出力内容を選んでください')); }

    return wb.xlsx.writeBuffer().then(function (buf) {
      return new Blob([buf], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
    });
  }

  // 公開 API（js/app.js から使う）。部品の入れ物 _ はそのまま残す
  var api = window.SKUKIKKSExport;
  api.build = build;
  api.collect = collect;
  api.parseDocs = parseDocs;
  api.wareki = wareki;
  api.KOSHU = KOSHU;
  api.loadHonbun = loadHonbun;
  api.honbunIsSogo = honbunIsSogo;
  api.sekoChapters = X.sekoChapters;
})(window.SKUKIKKSExport._);
