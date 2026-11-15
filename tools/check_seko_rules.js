// ②工事種目 → ③施工チェックシートの工事 の判定規則（js/app.js の SEKO_RULES ほか）の検査
// 使い方: node tools/check_seko_rules.js
// - SEKO_RULES・KOSHU_ITEMS の項目名・細目名が管理表データ（data/kanrihyo.js）に実在するか
// - 工種ごとに ②を選び直したとき（⑴は全選択）の ③の工事が EXPECT と同じか
// 問題が無ければ最後に OK と出す。管理表を差し替えたとき・規則を変えたときに実行する
var fs = require('fs');
var path = require('path');
var ROOT = path.join(__dirname, '..');
global.window = {};
require(path.join(ROOT, 'data', 'kanrihyo.js'));
require(path.join(ROOT, 'data', 'sekocheck.js'));
var DATA = window.KANRI_DATA;
var src = fs.readFileSync(path.join(ROOT, 'js', 'app.js'), 'utf8');

// app.js から必要な宣言だけを取り出して評価する（画面の部品は使わない）
function cut(start, end) {
  var a = src.indexOf(start);
  if (a < 0) { throw new Error('app.js に見つからない: ' + start); }
  var b = src.indexOf(end, a);
  return src.slice(a, b + end.length);
}
var state = { sel: {}, sekoChap: {} };
var ui = { seko: {} };
function itemCount(it) { var n = 0; it.groups.forEach(function (g) { if (state.sel[g.id]) { n++; } }); return n; }
function sekoChapters() { return window.SEKOCHECK_DATA.chapters.map(function (c) { return { id: c.id, name: c.name }; }); }
/* eslint-disable no-eval */
eval(cut('var COMMON =', ';'));
eval(cut('var T2 =', ';'));
eval(cut('var T3 =', ';'));
eval(cut('var T4 =', ';'));
eval(cut('var KOSHU_ITEMS =', '\n  };'));
eval(cut('var SEKO_RULES =', '];'));
eval(cut('var SEKO_CHAP =', '};'));
eval(cut('var sekoIndex =', ';'));
eval(cut('function sekoRuleIndex(', '\n  }\n'));
eval(cut('function sekoTargets(', '\n  }\n'));
eval(cut('function syncSekoChap(', '\n  }\n'));

// 工種（KOSHU の添字）ごとの ③の工事。KOSHU_ITEMS に登録のある工種だけ
var EXPECT = {
  0: 'ippan haikan', 1: 'ippan haisen', 2: 'ippan haikan haisen',
  3: 'ippan kiki', 4: 'ippan kiki', 5: 'ippan kiki', 6: 'ippan kiki',
  7: 'ippan juhen', 8: 'ippan juhen', 9: 'ippan juhen',
  10: 'ippan setchi', 11: 'ippan kakuu chichu'
};

var ng = 0;
function findItem(chId, name) {
  var ch = DATA.chapters.filter(function (c) { return c.id === chId; })[0];
  return ch ? ch.items.filter(function (i) { return i.name === name; }) : [];
}
SEKO_RULES.forEach(function (r) {
  var its = findItem(r[0], r[1]);
  if (!its.length) { console.log('NG SEKO_RULES: 項目が無い', r[0], r[1]); ng++; return; }
  (r[2] || []).forEach(function (e) {
    var ok = its.some(function (i) { return i.groups.some(function (g) { return g.e === e; }); });
    if (!ok) { console.log('NG SEKO_RULES: 細目が無い', r[1], e); ng++; }
  });
  r[3].forEach(function (id) {
    if (!sekoChapters().some(function (c) { return c.id === id; })) { console.log('NG SEKO_RULES: 工事id が無い', id); ng++; }
  });
});
Object.keys(SEKO_CHAP).forEach(function (k) {
  if (!sekoChapters().some(function (c) { return c.id === SEKO_CHAP[k]; })) { console.log('NG SEKO_CHAP: 工事id が無い', SEKO_CHAP[k]); ng++; }
});
Object.keys(KOSHU_ITEMS).forEach(function (k) {
  var map = KOSHU_ITEMS[k];
  Object.keys(map).forEach(function (chId) {
    map[chId].forEach(function (name) {
      if (!findItem(Number(chId), name).length) { console.log('NG KOSHU_ITEMS: 項目が無い', k, chId, name); ng++; }
    });
  });
});

Object.keys(KOSHU_ITEMS).forEach(function (k) {
  state.sel = {};
  DATA.chapters.forEach(function (ch) {
    var names = ch.id === 1 ? null : (KOSHU_ITEMS[k][ch.id] || []);
    ch.items.forEach(function (it) {
      if (names === null || names.indexOf(it.name) >= 0) { it.groups.forEach(function (g) { state.sel[g.id] = true; }); }
    });
  });
  syncSekoChap();
  var got = sekoChapters().filter(function (c) { return state.sekoChap[c.id]; }).map(function (c) { return c.id; }).join(' ');
  var mark = got === EXPECT[k] ? 'ok' : 'NG';
  if (mark === 'NG') { ng++; }
  console.log(mark, '工種' + k, got, mark === 'NG' ? '（期待: ' + EXPECT[k] + '）' : '');
});

// 何も選ばないときは0件、共通項目だけのときも0件
state.sel = {};
syncSekoChap();
if (Object.keys(state.sekoChap).some(function (k) { return state.sekoChap[k]; })) { console.log('NG 未選択で工事が出る'); ng++; }
console.log(ng ? 'NG ' + ng + ' 件' : 'OK');
process.exit(ng ? 1 : 0);
