# 選択した文字の字形と異体字を切り換える

[![Direct](https://img.shields.io/badge/Direct%20Link-ReplaceAlternateGlyphs.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/ReplaceAlternateGlyphs.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReplaceAlternateGlyphs.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択した文字の字形（JIS78字形・JIS90字形・旧字体など）を、ダイアログボックスで選んでその場で切り換える。
- 「高」→「髙」のような異体字（別の文字）への置き換えや、異体字と符号位置を並べた一覧の作成もできる。

### 主な機能

- 「字形」（左）
  - 標準字形・JIS78字形・JIS83字形・JIS90字形・JIS04字形・旧字体・エキスパート字形・等幅全角字形・等幅半角字形・等幅三分字形・等幅四分字形
  - 開いたときは、選択した文字に共通の字形が選ばれている。クリックするとすぐに反映する
- 「異体字」（右）
  - 選択の中に、登録された異体字のある文字があれば候補を「髙  U+9AD9」の形で並べる。クリックすると、選択内の同じ文字をすべてすぐに置き換える
  - 異体字のある文字が複数あるときは「高 → 髙  U+9AD9」のように元の文字を添える
  - 候補には元の文字も含むので、選び直せば戻せる
- ［一覧を作成］（左下）
  - レイヤー「_glyph」（無ければ最前面に作る）に、異体字の大きな字と符号位置を並べた一覧を作る。置き場所は元のテキストの右、フォントは元の文字のもの
- ［キャンセル］で、途中で反映した字形と異体字を元に戻す
- 日本語／英語の自動切り替え

### 使い方

1. 文字ツールで文字を選択する（テキストオブジェクトを選択してもよい）
2. スクリプトを実行する
3. 字形のラジオボタン、または異体字の候補をクリックする
4. ［OK］をクリックする

### オプション

スクリプト冒頭の「ユーザー設定」で変えられる。

- `ITAIJI_GROUPS`：異体字の組。1つの文字列に互いに異体字の文字を並べる（例 `"辺邊邉"`）。行を足すだけで増やせ、同じ文字を含む組は実行時にまとめる
- `GLYPH_LIST_LAYER_NAME`：一覧を作るレイヤー名（既定 `_glyph`）
- `GLYPH_LIST_GLYPH_SIZE` / `GLYPH_LIST_CODE_SIZE`：一覧の字と符号位置の大きさ（pt）

### 注意点

- 字形はフォントが持っていなければ見た目は変わらない
- ［一覧を作成］で作った一覧は、［キャンセル］しても残る
- 互換漢字（例 U+FA45）を含む異体字は、フォントによっては表示できない
- 空白・改行は対象にしない。連結したテキストはストーリー全体をまとめて扱い、ロック中・非表示のテキストは対象にしない

### 更新履歴

- v1.0.0（2026-10-17）初版
