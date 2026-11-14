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
  var OPT_KEYS = ['cover', 'gaiyo', 'soshiki', 'tenpu', 'kanri', 'check', 'hannyu', 'jishu', 'shiunten', 'hokoku', 'satsuei', 'list', 'docs', 'photos', 'page'];
  // 章内で最初の項目を選んだときに一緒に選ぶ共通項目
  var COMMON = ['共通事項', '機材の試験', '施工の立会い及び試験', '機器及び完成設備'];

  // 書類の種類ごとの出力内容の初期値（切り替えたときに設定する。あとから個別に変更できる）
  var DOC_TYPES = {
    sogo: {
      title: '総合施工計画書',
      hint: '表紙・工事概要・組織表・添付シートを出力します。管理表は含めません（工事種目の選択は不要）。',
      opt: { cover: true, gaiyo: true, soshiki: true, tenpu: true, kanri: false, check: false, list: false, docs: false, photos: false, hannyu: false, jishu: false, shiunten: false, hokoku: false, satsuei: false }
    },
    koshu: {
      title: '',
      hint: '工種ごとに作成します。表紙と、選択した工事種目の管理表（品質計画）・資料一覧・写真一覧を出力します。',
      opt: { cover: true, gaiyo: false, soshiki: false, tenpu: false, kanri: true, check: false, list: true, docs: true, photos: true, hannyu: false, jishu: false, shiunten: false, hokoku: false, satsuei: false }
    },
    check: {
      title: '',
      hint: '工種ごとの品質管理チェックシート（別添1）だけを、施工計画書とは別のファイルで出力します。',
      opt: { cover: false, gaiyo: false, soshiki: false, tenpu: false, kanri: false, check: true, list: false, docs: false, photos: false, hannyu: false, jishu: false, shiunten: false, hokoku: false, satsuei: false }
    },
    satsuei: {
      title: '',
      hint: '工種ごとの工事写真の撮影箇所（別添2）だけを、施工計画書とは別のファイルで出力します。',
      opt: { cover: false, gaiyo: false, soshiki: false, tenpu: false, kanri: false, check: false, list: false, docs: false, photos: false, hannyu: false, jishu: false, shiunten: false, hokoku: false, satsuei: true }
    },
    kan: {
      title: '施工計画書',
      hint: '簡略的な工事向け。表紙・工事概要・組織表と管理表を1冊にまとめて出力します。',
      opt: { cover: true, gaiyo: true, soshiki: true, tenpu: false, kanri: true, check: false, list: true, docs: true, photos: true, hannyu: false, jishu: false, shiunten: false, hokoku: false, satsuei: false }
    }
  };

  // 工種（SekokanExport.KOSHU の添字）ごとに、管理表で初期選択する項目。{ 章id: [項目名, ...] }
  // 工種別を選んだとき・工種を切り替えたときに、⑵〜⑼ の選択をこの内容に置き換える（⑴ 一般共通事項は変えない）。
  // 登録の無い工種（検査・試運転の計画書、その他）は選択を変えない
  var T2 = ['機材の試験', '共通事項', '施工の立会い及び試験'];
  var T3 = ['機材の試験', '共通事項', '施工の立会い及び試験', '機器及び完成設備'];
  var T4 = ['機材の試験', '据付け', '施工の立会い及び試験', '機器及び完成設備'];
  var KOSHU_ITEMS = {
    0: { 2: T2.concat(['電線保護物類', '金属管配線', '合成樹脂管配線（PF管、CD管、硬質ビニル管（通信・情報のみ））',
      '合成樹脂管配線（硬質ビニル管）', '金属製可とう電線管配線']) },
    1: { 2: T2.concat(['電線類', 'ケーブル配線（光ファイバケーブルを除く）', '光ファイバケーブル配線', '床上配線']) },
    2: { 2: T2.concat(['電線類', '電線保護物類', '金属ダクト配線']) },
    3: { 3: T3.concat(['分電盤', 'OA盤', '制御盤', '電灯設備', '動力設備']) },
    4: { 3: T3.concat(['配線器具', '照明器具', '防災用照明器具', '照明制御装置', '電灯設備']) },
    5: { 7: ['配線器具', '機材の試験', '施工の立会い及び試験', '機器及び完成設備', '端子盤・機器収納ラック等',
      '構内情報通信網装置', '構内交換装置', '情報表示装置', '映像・音響装置', '拡声装置', '誘導支援装置',
      'テレビ共同受信装置', '監視カメラ装置', '構内情報通信網設備', '構内交換設備', '情報表示設備', '映像・音響設備',
      '拡声設備', '誘導支援設備', 'テレビ共同受信設備', '監視カメラ設備'] },
    6: { 7: ['機材の試験', '施工の立会い及び試験', '機器及び完成設備', '自動火災報知装置', '自動閉鎖装置（自動閉鎖機構）',
      '非常警報装置', '自動火災報知設備', '自動閉鎖設備（自動閉鎖機構）', '非常警報設備'] },
    7: { 4: T4.concat(['キュービクル式配電盤']), 5: T4.concat(['直流電源装置']) },
    8: { 6: ['機材の試験', '施工の立会い及び試験', '機器及び完成設備', 'ディーゼルエンジン発電装置',
      'ディーゼルエンジン発電設備、ガスエンジン発電設備、ガスタービン発電設備及びマイクロガスタービン発電設備の据付け'] },
    9: { 6: ['機材の試験', '施工の立会い及び試験', '機器及び完成設備', '太陽光発電装置', '太陽光発電設備の据付け'] },
    10: { 2: T2.concat(['接地']) },
    11: { 2: T2.concat(['電線類', '電線保護物類', '外線材料', '地中配線', '架空配線']) }
  };

  var state = { info: {}, sel: {}, opt: {}, infoOpen: true, docType: '', koshu: 0 };
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
      state.docType = DOC_TYPES[saved.docType] ? saved.docType : '';
      state.koshu = Number(saved.koshu) || 0;
      if (!(state.koshu >= 0 && state.koshu <= window.SekokanExport.KOSHU.length)) { state.koshu = 0; }
    } else {
      // 初回: ⑴一般共通事項はどの工事にも該当するため選択済みにする
      DATA.chapters[0].items.forEach(function (it) {
        it.groups.forEach(function (g) { state.sel[g.id] = true; });
      });
    }
    if (!state.info.docTitle) { state.info.docTitle = '施工計画書'; }
    if (!state.info.sakuseiBi) { state.info.sakuseiBi = todayIso(); }
    // 書類の種類が未保存（初回、または旧版からの引き継ぎ）のときは「施工計画書（簡略）」にする
    if (!state.docType) {
      state.docType = 'kan';
      Object.keys(DOC_TYPES.kan.opt).forEach(function (k) {
        // 旧版で保存済みのチェックはそのまま使い、未設定の項目だけ初期値にする
        if (state.opt[k] === undefined) { state.opt[k] = DOC_TYPES.kan.opt[k]; }
      });
    }
    if (state.opt.check === undefined) { state.opt.check = false; }
    if (state.opt.hannyu === undefined) { state.opt.hannyu = false; }
    if (state.opt.jishu === undefined) { state.opt.jishu = false; }
    if (state.opt.shiunten === undefined) { state.opt.shiunten = false; }
    if (state.opt.hokoku === undefined) { state.opt.hokoku = false; }
    if (state.opt.satsuei === undefined) { state.opt.satsuei = false; }
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

  // 工種に合わせて管理表の項目を選び直す。登録の無い工種は false を返し、選択を変えない
  function selectByKoshu(koshu) {
    var map = KOSHU_ITEMS[koshu];
    if (!map) { return false; }
    DATA.chapters.forEach(function (ch) {
      if (ch.id === 1) { return; }
      var names = map[ch.id] || [];
      ch.items.forEach(function (it) { setItem(it, names.indexOf(it.name) >= 0); });
    });
    return true;
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
    var needSel = state.opt.kanri !== false;   // 管理表を出力する書類は工事種目の選択が必要
    $('summary').textContent = totalItems
      ? (totalChaps + '工事種目・' + totalItems + '項目 を選択中（管理項目 ' + totalRows + '行）')
      : (needSel ? '工事種目が選択されていません' : '工事種目の選択なしで出力できます');
    $('exportBtn').disabled = needSel && totalItems === 0;
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

  /* ---------- 書類の種類 ---------- */

  function koshuList() { return window.SekokanExport.KOSHU || []; }

  // 画面の表示を state.docType に合わせる
  function showDocType() {
    var btns = $('docType').querySelectorAll('button');
    var i;
    for (i = 0; i < btns.length; i++) {
      var on = btns[i].getAttribute('data-type') === state.docType;
      btns[i].className = on ? 'seg-btn on' : 'seg-btn';
      btns[i].setAttribute('aria-pressed', on ? 'true' : 'false');
    }
    var hint = DOC_TYPES[state.docType].hint;
    var typeByKoshu = state.docType === 'koshu' || state.docType === 'check' || state.docType === 'satsuei';
    var withCheck = state.opt.check === true;
    // チェックシートを出力するときは、書類の種類にかかわらず工種を選べるようにする
    var withHannyu = state.opt.hannyu === true;
    var withSatsuei = state.opt.satsuei === true;
    var byKoshu = typeByKoshu || withCheck || withHannyu || withSatsuei;
    if (withSatsuei && !(window.SATSUEI_DATA && window.SATSUEI_DATA[state.koshu])) {
      hint += 'この工種の撮影箇所は作成例に無いため、撮影対象表（電気設備工事編）の全行を出力します。';
    }
    if (state.docType === 'koshu') {
      hint += KOSHU_ITEMS[state.koshu]
        ? '工種を選ぶと、管理表の項目（⑵〜⑼）をその工種に合わせて選び直します。あとから個別に変更できます。'
        : 'この工種は管理表の項目を自動では選びません。必要な項目をチェックしてください。';
    }
    if (withCheck && !typeByKoshu) {
      hint += 'チェックシートは、下の「工種」で選んだ工種の内容を出力します。';
    }
    if (withCheck && state.koshu !== 14 && !(window.CHECK_DATA && window.CHECK_DATA[state.koshu])) {
      hint += 'この工種のチェック項目は未登録のため、空欄の様式を出力します。';
    }
    if (withHannyu) {
      hint += '機器搬入計画書は、共通の部分だけ記入した雛形を出力します（機器名は「工種」が搬入・据付け工事のときに入ります）。ワイヤロープの選定は、Excel で条件の欄を書き換えると再計算されます。';
    }
    if (state.opt.jishu === true) {
      hint += '使用前自主検査の検査項目は、検査項目ごとに1シートずつ出力します。B種接地抵抗許容値は、Excel で色付きの欄を書き換えると再計算されます。';
    }
    $('docTypeHint').textContent = hint;
    $('koshuField').hidden = !byKoshu;
  }

  // 書類の種類を切り替え、書類名と出力内容を初期値にする
  function applyDocType(type) {
    var def = DOC_TYPES[type];
    state.docType = type;
    Object.keys(def.opt).forEach(function (k) {
      state.opt[k] = def.opt[k];
      $('o_' + k).checked = def.opt[k];
    });
    var title = def.title;
    var list = koshuList();
    if (type === 'koshu') {
      title = state.koshu < list.length ? list[state.koshu] : '工種別施工計画書';
      selectByKoshu(state.koshu);
    }
    if (type === 'check') {
      title = (state.koshu < list.length ? list[state.koshu].replace(/施工計画書$|計画書$/, '') + ' ' : '')
        + '品質管理チェックシート';
    }
    if (type === 'satsuei') {
      title = (state.koshu < list.length ? list[state.koshu].replace(/施工計画書$|計画書$/, '') + ' ' : '')
        + '工事写真の撮影箇所';
    }
    state.info.docTitle = title;
    $('f_docTitle').value = title;
    showDocType();
    changed();
  }

  function bindDocType() {
    var sel = $('f_koshu');
    koshuList().forEach(function (name, i) {
      var o = el('option', null, name);
      o.value = String(i);
      sel.appendChild(o);
    });
    var other = el('option', null, 'その他（書類名を入力）');
    other.value = String(koshuList().length);
    sel.appendChild(other);
    sel.value = String(state.koshu);
    sel.addEventListener('change', function () {
      state.koshu = Number(sel.value) || 0;
      if (state.docType === 'koshu' || state.docType === 'check' || state.docType === 'satsuei') {
        applyDocType(state.docType);       // 書類名を工種に合わせる
      } else {
        showDocType();                     // 施工計画書のみ・総合では、書類名や出力内容は変えない
        changed();
      }
    });
    var btns = $('docType').querySelectorAll('button');
    var i;
    function bind(b) {
      b.addEventListener('click', function () {
        var type = b.getAttribute('data-type');
        // 選択中の種類をもう一度押しても、入力済みの書類名や出力内容は初期値に戻さない
        if (type === state.docType) { return; }
        applyDocType(type);
      });
    }
    for (i = 0; i < btns.length; i++) { bind(btns[i]); }
    showDocType();
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
        showDocType();
        refresh();
        save();
        hideResult();
      });
    });
    bindDocType();
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
