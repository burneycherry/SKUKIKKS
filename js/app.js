/* 施工計画書メーカー 画面制御
 * 依存: data/kanrihyo.js (window.KANRI_DATA), js/export.js (window.SekokanExport)
 */
(function () {
  'use strict';

  var DATA = window.KANRI_DATA;
  var STORE_KEY = 'sekokan.v1';
  var INFO_KEYS = ['docTitle', 'kojiName', 'basho', 'kokiFrom', 'kokiTo', 'hacchu',
    'juchu', 'dairinin', 'gijutsu', 'sakuseisha', 'sakuseiBi',
    'kojiKanri', 'kanri', 'sekkei', 'yoto', 'chiiki', 'kozo', 'kibo', 'shobo', 'bunrui'];
  var OPT_KEYS = ['cover', 'gaiyo', 'soshiki', 'list', 'docs', 'photos', 'page'];
  // 章内で最初の項目を選んだときに一緒に選ぶ共通項目
  var COMMON = ['共通事項', '機材の試験', '施工の立会い及び試験', '機器及び完成設備'];

  var state = { info: {}, sel: {}, opt: {}, infoOpen: true };
  var ui = { chap: {}, item: {}, group: {} };
  var lastFile = null;   // { blob, name, url }

  function $(id) { return document.getElementById(id); }

  function el(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) { e.className = cls; }
    if (text !== undefined && text !== null) { e.textContent = text; }
    return e;
  }

  /* ---------- 保存・復元 ---------- */

  function todayIso() {
    var d = new Date();
    function p(n) { return (n < 10 ? '0' : '') + n; }
    return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
  }

  function load() {
    var saved = null;
    try {
      var raw = window.localStorage.getItem(STORE_KEY);
      if (raw) { saved = JSON.parse(raw); }
    } catch (e) { saved = null; }
    if (saved && typeof saved === 'object') {
      state.info = saved.info || {};
      state.sel = saved.sel || {};
      state.opt = saved.opt || {};
      state.infoOpen = saved.infoOpen !== false;
    } else {
      // 初回: ⑴一般共通事項はどの工事にも該当するため選択済みにする
      DATA.chapters[0].items.forEach(function (it) {
        it.groups.forEach(function (g) { state.sel[g.id] = true; });
      });
    }
    if (!state.info.docTitle) { state.info.docTitle = '施工計画書'; }
    if (!state.info.sakuseiBi) { state.info.sakuseiBi = todayIso(); }
    OPT_KEYS.forEach(function (k) {
      if (state.opt[k] === undefined) { state.opt[k] = true; }
    });
  }

  function save() {
    try {
      window.localStorage.setItem(STORE_KEY, JSON.stringify(state));
    } catch (e) { /* 保存できない環境でも動作は継続 */ }
  }

  /* ---------- 選択状態 ---------- */

  function itemCount(it) {
    var n = 0;
    it.groups.forEach(function (g) { if (state.sel[g.id]) { n++; } });
    return n;
  }

  function setItem(it, on) {
    it.groups.forEach(function (g) {
      if (on) { state.sel[g.id] = true; } else { delete state.sel[g.id]; }
    });
  }

  function chapterHasSelection(ch) {
    return ch.items.some(function (it) { return itemCount(it) > 0; });
  }

  function addCommon(ch) {
    if (ch.id === 1) { return; }
    ch.items.forEach(function (it) {
      if (COMMON.indexOf(it.name) >= 0) { setItem(it, true); }
    });
  }

  function refresh() {
    var totalItems = 0;
    var totalRows = 0;
    var totalChaps = 0;
    DATA.chapters.forEach(function (ch) {
      var u = ui.chap[ch.id];
      var on = 0;
      var part = false;
      ch.items.forEach(function (it) {
        var n = itemCount(it);
        var iu = ui.item[it.id];
        iu.cb.checked = n > 0 && n === it.groups.length;
        iu.cb.indeterminate = n > 0 && n < it.groups.length;
        if (n > 0) { on++; }
        if (n > 0 && n < it.groups.length) { part = true; }
        it.groups.forEach(function (g) {
          var gcb = ui.group[g.id];
          if (gcb) { gcb.checked = !!state.sel[g.id]; }
          if (state.sel[g.id]) { totalRows += Math.max(1, g.rows.length); }
        });
      });
      u.cb.checked = on === ch.items.length && !part;
      u.cb.indeterminate = on > 0 && !u.cb.checked;
      u.count.textContent = on + '/' + ch.items.length;
      u.count.className = on > 0 ? 'count on' : 'count';
      totalItems += on;
      if (on > 0) { totalChaps++; }
    });
    $('summary').textContent = totalItems
      ? (totalChaps + '工事種目・' + totalItems + '項目 を選択中（管理項目 ' + totalRows + '行）')
      : '工事種目が選択されていません';
    $('exportBtn').disabled = totalItems === 0;
  }

  function changed() {
    refresh();
    save();
    hideResult();
  }

  /* ---------- 工事種目リストの描画 ---------- */

  function renderGroups(it, box) {
    it.groups.forEach(function (g) {
      var lab = el('label', 'group');
      var cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.checked = !!state.sel[g.id];
      cb.addEventListener('change', function () {
        if (cb.checked) { state.sel[g.id] = true; } else { delete state.sel[g.id]; }
        changed();
      });
      var sp = el('span', null, g.e || '（細目なし）');
      if (g.c) { sp.appendChild(el('small', null, g.c.replace(/\n/g, ' '))); }
      lab.appendChild(cb);
      lab.appendChild(sp);
      box.appendChild(lab);
      ui.group[g.id] = cb;
    });
  }

  function renderItem(ch, it) {
    var wrap = el('div', 'item');
    var row = el('div', 'item-row');
    var lab = el('label', 'item-label');
    var cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.addEventListener('change', function () {
      var first = cb.checked && !chapterHasSelection(ch);
      setItem(it, cb.checked);
      if (first) { addCommon(ch); }
      changed();
    });
    var text = el('span', 'item-text');
    text.appendChild(el('span', 'item-name', it.name));
    var names = [];
    it.groups.forEach(function (g) {
      if (g.e && names.indexOf(g.e) < 0) { names.push(g.e); }
    });
    var sub = names.slice(0, 4).join('、') + (names.length > 4 ? ' ほか' : '');
    text.appendChild(el('span', 'item-sub', sub));
    lab.appendChild(cb);
    lab.appendChild(text);
    row.appendChild(lab);

    var more = el('button', 'item-more');
    more.type = 'button';
    more.setAttribute('aria-expanded', 'false');
    more.appendChild(el('span', null, '細目 ' + it.groups.length));
    more.appendChild(el('span', 'chev'));
    var box = el('div', 'groups');
    box.hidden = true;
    var built = false;
    more.addEventListener('click', function () {
      if (!built) { renderGroups(it, box); built = true; }
      box.hidden = !box.hidden;
      more.setAttribute('aria-expanded', box.hidden ? 'false' : 'true');
    });
    row.appendChild(more);
    wrap.appendChild(row);
    wrap.appendChild(box);

    var search = (it.name + ' ' + it.kubun + ' ' + names.join(' ')).toLowerCase();
    ui.item[it.id] = { cb: cb, el: wrap, search: search };
    return wrap;
  }

  function renderChapters() {
    var root = $('chapters');
    DATA.chapters.forEach(function (ch) {
      var box = el('div', 'chap');
      var head = el('div', 'chap-head');
      var cbWrap = el('label', 'chap-cb');
      var cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.setAttribute('aria-label', ch.name + ' をすべて選択');
      cb.addEventListener('change', function () {
        var on = cb.checked;
        ch.items.forEach(function (it) { setItem(it, on); });
        changed();
      });
      cbWrap.appendChild(cb);
      var btn = el('button', 'chap-btn');
      btn.type = 'button';
      btn.setAttribute('aria-expanded', 'false');
      btn.appendChild(el('span', 'chap-name', ch.title.replace(/　/g, ' ')));
      var count = el('span', 'count', '');
      btn.appendChild(count);
      btn.appendChild(el('span', 'chev'));
      head.appendChild(cbWrap);
      head.appendChild(btn);

      var body = el('div', 'chap-body');
      body.hidden = true;
      var kubun = null;
      var kubunEls = [];
      ch.items.forEach(function (it) {
        if (it.kubun !== kubun) {
          kubun = it.kubun;
          var k = el('div', 'kubun', kubun);
          kubunEls.push({ el: k, items: [] });
          body.appendChild(k);
        }
        var node = renderItem(ch, it);
        kubunEls[kubunEls.length - 1].items.push(it.id);
        body.appendChild(node);
      });
      btn.addEventListener('click', function () {
        body.hidden = !body.hidden;
        btn.setAttribute('aria-expanded', body.hidden ? 'false' : 'true');
      });
      box.appendChild(head);
      box.appendChild(body);
      root.appendChild(box);
      ui.chap[ch.id] = { cb: cb, count: count, el: box, body: body, btn: btn, kubun: kubunEls };
    });
  }

  function applyFilter() {
    var q = $('filter').value.replace(/^\s+|\s+$/g, '').toLowerCase();
    var any = false;
    DATA.chapters.forEach(function (ch) {
      var u = ui.chap[ch.id];
      var hit = 0;
      u.kubun.forEach(function (k) {
        var khit = 0;
        k.items.forEach(function (id) {
          var iu = ui.item[id];
          var show = !q || iu.search.indexOf(q) >= 0;
          iu.el.hidden = !show;
          if (show) { khit++; }
        });
        k.el.hidden = khit === 0;
        hit += khit;
      });
      u.el.hidden = hit === 0;
      if (q && hit > 0) {
        u.body.hidden = false;
        u.btn.setAttribute('aria-expanded', 'true');
      }
      if (hit > 0) { any = true; }
    });
    $('noHit').hidden = any;
  }

  /* ---------- 入力欄 ---------- */

  function bindInputs() {
    INFO_KEYS.forEach(function (k) {
      var inp = $('f_' + k);
      inp.value = state.info[k] || '';
      inp.addEventListener('input', function () {
        state.info[k] = inp.value;
        save();
        hideResult();
      });
    });
    OPT_KEYS.forEach(function (k) {
      var cb = $('o_' + k);
      cb.checked = state.opt[k] !== false;
      cb.addEventListener('change', function () {
        state.opt[k] = cb.checked;
        save();
        hideResult();
      });
    });
    var toggle = $('infoToggle');
    var body = $('infoBody');
    function setOpen(open) {
      body.hidden = !open;
      toggle.setAttribute('aria-expanded', open ? 'true' : 'false');
    }
    setOpen(state.infoOpen);
    toggle.addEventListener('click', function () {
      state.infoOpen = body.hidden;
      setOpen(state.infoOpen);
      save();
    });
    $('filter').addEventListener('input', applyFilter);
    $('clearAll').addEventListener('click', function () {
      if (!window.confirm('すべてのチェックを外しますか？')) { return; }
      state.sel = {};
      changed();
    });
  }

  /* ---------- Excel 出力 ---------- */

  function fileName() {
    var base = (state.info.docTitle || '施工計画書');
    if (state.info.kojiName) { base += '_' + state.info.kojiName; }
    base = base.replace(/[\\\/:*?"<>|\s]+/g, '_').substring(0, 80);
    return base + '_' + (state.info.sakuseiBi || todayIso()).replace(/-/g, '') + '.xlsx';
  }

  function hideResult() {
    $('result').hidden = true;
    $('error').hidden = true;
    if (lastFile && lastFile.url) {
      try { URL.revokeObjectURL(lastFile.url); } catch (e) { /* 無視 */ }
    }
    lastFile = null;
  }

  function showError(msg) {
    var e = $('error');
    e.textContent = msg;
    e.hidden = false;
  }

  function canShareFile(file) {
    try {
      return !!(navigator.canShare && navigator.canShare({ files: [file] }));
    } catch (e) { return false; }
  }

  function doExport() {
    var btn = $('exportBtn');
    hideResult();
    btn.disabled = true;
    btn.textContent = '作成中…';
    window.SekokanExport.build(DATA, state).then(function (blob) {
      var name = fileName();
      var url = URL.createObjectURL(blob);
      lastFile = { blob: blob, name: name, url: url };
      var link = $('dlLink');
      link.href = url;
      link.setAttribute('download', name);
      $('resultName').textContent = '作成しました：' + name + '（' + Math.round(blob.size / 1024) + ' KB）';
      var file = null;
      try { file = new File([blob], name, { type: blob.type }); } catch (e) { file = null; }
      lastFile.file = file;
      $('shareBtn').hidden = !(file && canShareFile(file));
      $('result').hidden = false;
      $('result').scrollIntoView({ block: 'center' });
      // PC等ではそのまま保存を開始する (iOSのホーム画面アプリでは下のボタンを使用)
      try { link.click(); } catch (e) { /* ボタンから再実行可能 */ }
    }).catch(function (err) {
      showError('Excelの作成に失敗しました：' + (err && err.message ? err.message : err));
    }).then(function () {
      btn.textContent = 'Excel出力';
      refresh();
    });
  }

  function doShare() {
    if (!lastFile || !lastFile.file) { return; }
    navigator.share({ files: [lastFile.file], title: lastFile.name }).catch(function (err) {
      if (err && err.name === 'AbortError') { return; }
      showError('共有できませんでした。「ダウンロード」をお試しください。');
    });
  }

  /* ---------- 起動 ---------- */

  function init() {
    if (!DATA || !DATA.chapters) {
      showError('管理表データを読み込めませんでした。');
      return;
    }
    load();
    renderChapters();
    bindInputs();
    refresh();
    $('exportBtn').addEventListener('click', doExport);
    $('shareBtn').addEventListener('click', doShare);
    if ('serviceWorker' in navigator && /^https?:$/.test(location.protocol)) {
      navigator.serviceWorker.register('sw.js').catch(function () { /* オフライン機能なしで継続 */ });
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
