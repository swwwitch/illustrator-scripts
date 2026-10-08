# 指定したファイルを開いて PDF として保存し、閉じる

[![Direct](https://img.shields.io/badge/Direct%20Link-OpenSaveAsPDFAndClose.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/export/OpenSaveAsPDFAndClose.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/OpenSaveAsPDFAndClose.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選んだファイルを Illustrator で1つずつ開き、ダイアログボックスで選んだ PDF プリセット・裁ち落とし・トンボで PDF として保存して閉じます。

Illustrator 付属のスクリプト「ドキュメントを PDF として保存.jsx」をもとにしています。付属版は開いているドキュメントが対象ですが、こちらは開いていないファイルを対象にします。

### 主な機能

- 対象ファイル（.ai / .eps / .pdf / .svg）を複数まとめて選択
- PDF プリセットを選択
- 裁ち落とし（mm、上下左右同じ幅）とトンボ（日本式／西洋式）を付ける
- 保存先を、元のファイルと同じ場所か指定したフォルダーから選択
- 同名の PDF があるときは、上書き・連番を付ける・スキップから選択
- フォントやリンクの警告で止まらずに処理
- 前回の設定を記憶
- 日本語／英語UI

### 使い方

1. `OpenSaveAsPDFAndClose.jsx` を実行します。
2. ［ファイル］パネルの［選択...］で対象のファイルを選びます。
3. PDF の設定と保存先を選びます。
4. ［保存］をクリックすると、ファイルを順に開いて PDF に保存し、閉じます。終わると、保存した件数とスキップ・失敗したファイルを表示します。

PDF のファイル名は、元のファイル名の拡張子を `.pdf` に替えたものです（例: `sample.ai` → `sample.pdf`）。

### ［ファイル］パネル

| 項目 | 内容 |
| --- | --- |
| 対象 | 選んだファイルの件数。マウスを重ねるとファイル名の一覧を表示します。 |
| ［選択...］ | 対象のファイルを選びます。選び直すと、前の選択と置き換わります。 |

### ［PDF］パネル

| 項目 | 内容 |
| --- | --- |
| PDF プリセット | 保存に使う PDF プリセット。初期値は［Illustrator 初期設定］です。 |
| 裁ち落とし | 上下左右に同じ幅の裁ち落としを付けます（単位は mm）。ドキュメントの裁ち落とし設定は使いません。 |
| トンボ | トンボを付けます。日本式／西洋式を選べます。 |

### ［保存先］パネル

| 項目 | 内容 |
| --- | --- |
| 元のファイルと同じ場所 | 元のファイルと同じフォルダーに保存します。 |
| 指定 | ［選択...］で選んだフォルダーに保存します。フォルダーが無いときは作ります。 |
| 同名の PDF を上書き | 同名の PDF があるときに上書きします。 |
| 上書きしないとき | ［連番を付ける］は「名前-1.pdf」「名前-2.pdf」… の空いている名前で保存します。［スキップ］は保存しません。 |

### 設定変数

スクリプト冒頭の「ユーザー設定」で初期値を変更できます。

| 変数 | 初期値 | 内容 |
| --- | --- | --- |
| `TARGET_FILES` | `[]` | 最初に対象にするファイルのパス。空なら［選択...］で選びます |
| `OPENABLE_FILE_PATTERN` | `/\.(ai\|eps\|pdf\|svg)$/i` | ［選択...］で選べるファイル |
| `DEFAULT_PDF_PRESETS` | `["[Illustrator Default]", "[Illustrator 初期設定]"]` | PDF プリセットの初期値の候補（上から順に、あるものを使う） |
| `DEFAULT_SETTINGS` | — | ダイアログボックスの初期値（2回目からは前回の設定） |

### 注意点

- 元の PDF 自身と、同じ実行で保存した PDF は、［同名の PDF を上書き］がオンでも上書きしません（［上書きしないとき］の設定に従います）。
- すでに開いているファイルはスキップします（保存すると編集中のドキュメントが PDF に切り替わるため）。
- 裁ち落としとトンボは、オフにしても PDF プリセット側の設定ではなく「なし」で保存します。
- PDF を開いたときは1ページ目だけが開くため、複数ページの PDF を対象にすると、2ページ目以降は保存した PDF に入りません。
- 設定は `Folder.userData` の `illustrator-scripts/OpenSaveAsPDFAndClose.json` に保存します。

### 更新履歴

- v1.1.1（2026-10-08）元の PDF 自身や、同じ実行で保存した PDF を上書きしないようにした（［同名の PDF を上書き］がオンでも連番を付けるかスキップ）
- v1.1.0（2026-10-08）初期バージョン
