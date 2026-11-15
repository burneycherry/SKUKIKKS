// 出力の回帰テスト: 全工種・全シートの Excel を生成して保存する
// 使い方: python3 -m http.server 8765 を起動してから
//   node tools/test_export.js <出力先>        … 変更前と変更後で2回実行し、
//   python3 tools/compare_xlsx.py <変更前> <変更後> で比較する
// 引数3 に URL を渡すと、そのサーバーで実行する（既定 http://localhost:8765/index.html）
const { chromium } = require('playwright');
const fs = require('fs');
(async () => {
  const out = process.argv[2];
  if (!out) { console.error('出力先のディレクトリを指定してください'); process.exit(1); } fs.mkdirSync(out, { recursive: true });
  const b = await chromium.launch(); const p = await b.newPage();
  const errs = []; p.on('pageerror', e => errs.push(e.message));
  await p.goto(process.argv[3] || 'http://localhost:8765/index.html');
  const res = await p.evaluate(async () => {
    const all = {}; KANRI_DATA.chapters.forEach(c => c.items.forEach(i => i.groups.forEach(g => { all[g.id] = true; })));
    const info = { docTitle: '試験書類', kojiName: '試験工事', basho: '東京都', kokiFrom: '2026-04-01', kokiTo: '2027-03-31',
      hacchu: '発注者', juchu: '受注者', dairinin: '代理 人', gijutsu: '技術 者', sakuseisha: '作成 者', sakuseiBi: '2026-11-15',
      kojiKanri: '監理部', kanri: '設計事務所', sekkei: '設計者', yoto: '庁舎', kozo: 'RC', chiiki: '商業', kibo: '5階', shobo: '15項', bunrui: '一般' };
    const on = { cover: true, honbun: true, gaiyo: true, soshiki: true, tenpu: true, kanri: true, check: true, satsuei: true,
      hannyu: true, jishu: true, shiunten: true, hokoku: true, list: true, docs: true, photos: true, page: true };
    const cases = [];
    for (let k = 0; k <= 16; k++) cases.push(['koshu' + k, { docType: 'koshu', koshu: k, info, sel: all, opt: on }]);
    cases.push(['sogo', { docType: 'sogo', koshu: 0, info, sel: {}, opt: Object.assign({}, on, { kanri: false, list: false }) }]);
    cases.push(['seko', { docType: 'seko', koshu: 0, info, sel: {}, opt: { cover: false, gaiyo: false, soshiki: false, tenpu: false, kanri: false, list: false, docs: false, photos: false, seko: true, shinsei: true } }]);
    cases.push(['kan_nopage', { docType: 'kan', koshu: 3, info: {}, sel: all, opt: Object.assign({}, on, { page: false, honbun: false }) }]);
    const r = {};
    for (const [n, st] of cases) {
      const blob = await SKUKIKKSExport.build(KANRI_DATA, st);
      const buf = new Uint8Array(await blob.arrayBuffer());
      let s = ''; for (let i = 0; i < buf.length; i += 8192) s += String.fromCharCode.apply(null, buf.subarray(i, i + 8192));
      r[n] = btoa(s);
    }
    r._api = Object.keys(SKUKIKKSExport).sort().join(',') + '|' + SKUKIKKSExport.KOSHU.length + '|' + SKUKIKKSExport.wareki('2026-11-15');
    return r;
  });
  for (const n in res) { if (n === '_api') { fs.writeFileSync(out + '/_api.txt', res[n]); continue; } fs.writeFileSync(out + '/' + n + '.xlsx', Buffer.from(res[n], 'base64')); }
  console.log('files', Object.keys(res).length, 'errors', errs);
  if (errs.length) { process.exitCode = 1; }
  await b.close();
})();
