# CLAUDE.md — SKUKIKKS（施工計画書メーカー）

総合施工計画書・工種別施工計画書・施工計画書（簡略）を Excel(xlsx) 出力する静的 Web アプリ（PWA）。
工事種目をチェックすると、国交省の管理表から該当部分を抜粋して品質計画として付ける。
ビルド工程なし。`index.html` を開けば動く。回答・コメント・UI文言はすべて日本語。

## 対象環境

- 主対象: iOS Safari / ホーム画面に追加した PWA。Android Chrome・PC も動作可とする
- サーバー処理なし。入力値は localStorage（キー `SKUKIKKS.v1`。旧キー `sekokan.v1` があれば読み込んで引き継ぐ）のみに保存

## ファイル構成

| パス | 役割 |
|---|---|
| `index.html` | 画面（ヘッダーの「使い方」 ①工事情報 ②工事種目の選択 ③出力内容 ＋下部バー） |
| `css/style.css` | スタイル（ライト／ダーク） |
| `js/app.js` | 画面制御・選択状態・保存・ダウンロード／共有 |
| `js/export/*.js` | Excel ブック生成（`window.SKUKIKKSExport`）。書類ごとに9ファイル。分担は `docs/shiyo/excel.md` |
| `data/kanrihyo.js` | 管理表データ（自動生成。手で編集しない） |
| `data/honbun.js` | 本文シートのデータ（`window.HONBUN_DATA`。自動生成。手で編集しない） |
| `data/checklist.js` | 品質管理チェックシートのデータ（`window.CHECK_DATA`。工種の添字ごと。手で編集する） |
| `data/satsuei.js` | 工事写真の撮影箇所のデータ（`window.SATSUEI_DATA`。工種の添字ごと＋全行 `all`。手で編集する） |
| `tools/build_data.py` | `source/*.xlsx` → `data/kanrihyo.js` 変換（要 openpyxl） |
| `tools/build_honbun.py` | `docs/honsho` の作成例の記録 → `data/honbun.js` 変換（記録を直したら再実行） |
| `tools/test_export.js` / `compare_xlsx.py` | 出力の回帰テスト（全工種の Excel を作って変更前と比べる） |
| `source/` | 国交省 管理表（令和8年版）原本 xlsx 9ファイル |
| `vendor/` | ExcelJS 4.4.0（MIT。ライセンス同梱） |
| `sw.js` / `manifest.webmanifest` / `icons/` | PWA 関連 |
| `docs/shiyo/` | アプリの詳細仕様（書類の種類 `shorui.md`・画面 `gamen.md`・Excel 出力 `excel.md`） |
| `docs/honsho/` | 書籍「電気設備工事の施工管理」の読み取り記録（仕様検討の参照用）。確認・保留事項は `00_kakunin.md` |
| `docs/hyojunzu/` | 国交省「標準図（電気設備工事編）令和7年版」PDF の URL とページ索引（「電力28」等の番号との対応） |

## コーディング規約（必須）

- **CSS カスタムプロパティ（変数）は絶対に使わない**。色・寸法は直接値で書く
  （コピペ時に記号が変換されて壊れるため）。`env(safe-area-inset-*)` は使用可
- JavaScript は SyntaxError を避けるため保守的な構文で書く
  - `var` / `function` / Promise の `.then()` を基本とする
  - オプショナルチェーン、null 合体、クラスフィールド、トップレベル await は使わない
  - 文字列中の全角記号・引用符の混入に注意。編集後は `node --check` とブラウザで読み込み確認
- 外部 CDN は使わない（オフライン動作のため `vendor/` に同梱）
- 入力欄の文字サイズは 16px 以上（iOS の自動ズーム防止）、タップ領域は 44px 以上
- localStorage の読み書きは必ず try/catch で囲む

## Excel 出力のファイル分担（`js/export/`）

- 読み込み順は `index.html` のとおり: `core` → `sogo` → `kanri` → `honbun` → `hannyu` → `kensa` → `shiunten` → `check` → `main`
- 各ファイルは `(function (X) { … })(window.SKUKIKKSExport._);` の形。先頭で `var box = X.box, …` と部品を取り出し、
  末尾で `X.名前 = 名前;` と公開する。公開 API は `main.js` で付ける
- 2つ以上のファイルで使う部品は `core.js` に置く。シート順は `main.js` の `buildNow` で決まる
- ファイルを足したら `index.html` の `<script>`（`main.js` より前）と `sw.js` の `FILES` にも足す

## 仕様の要点（詳細は `docs/shiyo/`）

- 書類の種類は 総合 `sogo`／工種別 `koshu`／施工計画書のみ `kan`／チェックシート `check`／撮影箇所 `satsuei`
  （`js/app.js` の `DOC_TYPES`）。切り替えると書類名と出力内容が初期値になる → `shorui.md`
- 工種別は1工種ずつ作る。工種を選ぶと管理表の ⑵〜⑼ をその工種の項目に置き換える（`KOSHU_ITEMS`。⑴ は変えない）
- 機器搬入計画書ほか・使用前自主検査・総合試運転・報告書の表は「③出力内容」のチェックで出す（初期値オフ）
- 本文シートは作成例の文章（`data/honbun.js`。出力するときだけ読み込む）。原本から表記を変えたものは
  `tools/build_honbun.py` の `FIX`、自動で作れない項は `OVERRIDE`。除いた行は `--dropped` で確認する
- 表紙・工事概要・組織表ほかは書籍の総合施工計画書作成例にならう。見出し等の表記は作成要領側に合わせる
- 作成例の数値が「標仕」と違うときは「標仕」に合わせる
- 図や表は1つずつ独立したシートにする。全体工程表・総合仮設計画などは表題だけ出し、利用者が Excel で仕上げる
- 管理表は原本と同じ列構成（A区分〜N掲載ページ）。A4横・横幅1ページ。フォントは「ＭＳ Ｐゴシック」
- 行の高さは文字数から見積もって明示設定（iOS の Numbers 等でも崩れにくくするため）
- 日付は令和表記（`SKUKIKKSExport.wareki`）。出力ファイル名は `書類名_工事名_作成日(yyyymmdd).xlsx`
- iOS のホーム画面アプリでは download が効かないことがあるため `navigator.share` を併設（生成と共有は別タップ）

## 管理表を差し替えるとき

1. `source/R08_denki_kanrihyo_N.xlsx`（N=1〜9）を置き換える
2. `python3 tools/build_data.py` を実行し、項目数・行数の出力を確認
3. `sw.js` の `VERSION` を上げる
4. 細目 id が変わるため、端末に保存済みのチェック状態は引き継がれない点に注意

## 本書の記録（`docs/honsho/`）

- 記録ルールは `docs/honsho/README.md`。確認・保留事項は `00_kakunin.md` に K 番号（本書のページ順）で管理する
- 新しい確認事項は K-99 からの続き番号。旧番号との対応は `00_kakunin_kyuban.md`
- 「対象」には本書のページ（p.○）を必ず書く

## 動作確認

```bash
python3 -m http.server 8765
```

- `http://localhost:8765/` を開き、項目をチェック → Excel出力 → 出力ファイルを開いて確認
- Playwright（Chromium, iPhone エミュレーション）で確認可能。
  ヘッドレスでは日本語ファイル名が `download` と報告されるが実ブラウザでは正常
- 体裁確認は `soffice --headless --convert-to pdf` で PDF 化すると早い
- 出力を変えない修正（整理・分割など）は回帰テストで差が無いことを確かめる

```bash
node tools/test_export.js /tmp/before
```

```bash
python3 tools/compare_xlsx.py /tmp/before /tmp/after
```

  （変更前に1回目、変更後に `/tmp/after` へ2回目を実行。作成日時以外が完全一致なら `differences 0`）

## 変更時の注意

- 配信ファイルを変更したら `sw.js` の `VERSION` を必ず更新（古いキャッシュが残るため）
- ファイルを追加したら `sw.js` の `FILES` にも追加
- 仕様を変えたら `docs/shiyo/` の該当ファイルも直す
- CLAUDE.md は 100〜200 行に保ち、詳細は別ファイル（`docs/` 等）に分けて管理する
- 不明な点は推測で埋めず「不明」と明記する
