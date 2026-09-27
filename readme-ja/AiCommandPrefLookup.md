# メニューコマンド・ツール・環境設定のコードを引く

[![Direct](https://img.shields.io/badge/Direct%20Link-AiCommandPrefLookup.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/AiCommandPrefLookup.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiCommandPrefLookup.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

スクリプトに埋め込んだ一覧から項目を選んで、`app.executeMenuCommand()`・`app.selectTool()`・`app.preferences` の get／set のコードを出力します。メニュー名や設定の日本語名から ID・キーを引く辞典として使えます。

### 主な機能

- 種類：［メニューコマンド］［環境設定］［ツール］を切り替え
  - メニューコマンド：`app.executeMenuCommand('…');`
  - ツール：`app.selectTool('…');`
  - 環境設定：`app.preferences.get〜Preference('…');` と `set〜Preference('…', 値);` の2行（型に合わせて Boolean／Integer／Real／String を使い分け、値の例は実機の環境設定ファイルの値）
- 言語：メニュー名・設定名を［日本語］［英語］で切り替え
- カテゴリ：メニューの最上位（ファイル、編集、オブジェクト…）や環境設定の区分で絞り込み
- キーワード：名前と ID・キーを入力するたびに絞り込み（大文字小文字を区別しない、正規表現も可）
- 複数選択すると、リストの順にコードを並べて出力
- ［メニュー名をコメントで付ける］：各行の末尾に `// ファイル > 新規...` のようにメニュー名を付ける（初期値 OFF）
- メモ：環境設定キーの値の意味（`rulerType` の単位コードなど）や、反映に再起動が必要といった注意点を表示
- ［再調査…］：Illustrator のバージョンアップ後に、一覧の更新が必要な項目を調べる（下記）

### 使い方

1. スクリプトを実行します（ドキュメントを開いていなくても使えます）
2. 種類と言語を選び、カテゴリやキーワードで絞り込みます
3. リストから項目を選ぶと、名前・メモ・コードが表示されます
4. コード欄を全選択してコピーし、スクリプトに貼り付けます

### 再調査

［再調査…］を押すと、実行中の Illustrator の設定フォルダーを読み、一覧との差分を表示します。

- ショートカットファイル（.kys）と照合し、改名・廃止されたかもしれないメニューコマンドとツール、一覧にない ID を列挙
- 環境設定ファイル（Adobe Illustrator Cloud Prefs／Adobe Illustrator 環境設定）と照合し、一覧にないキーと型の食い違いを列挙
- 照合元の URL（Adobe Community のスレッド、Ai Command Palette、sttk3 氏の Notion データベース、Ten A 氏の一覧）をブラウザーで開く
- 設定フォルダーを開く、結果をテキストに書き出す

一覧を照合した Illustrator のバージョンと実行中のバージョンが違うと、結果の冒頭に注意書きが出ます。

### 注意点

- 一覧はスクリプトに埋め込んであります。元データは、メニューコマンドの一覧と環境設定キーの一覧をまとめたテキストで、更新したらスクリプトに埋め込み直します
- ID の末尾の空白も ID の一部です（例：`'Live PSAdapter_plugin_Ct  '`）。出力したコードの空白は消さないでください
- ショートカットファイルは、［キーボードショートカット］でセットを保存したときに作られます。無い場合、再調査ではメニューコマンドとツールを照合できません
- ショートカットファイルや環境設定ファイルに無いことは、廃止の根拠になりません。ショートカットの対象外のコマンドや、一度も変更していない設定はファイルに書かれないためです
- コードのコピーはコード欄を選択して行います（Illustrator のスクリプトからは文字列を直接クリップボードに送れないため）

### 更新履歴

- v1.0.0 (20260927) : 初期バージョン
