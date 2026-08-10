# ImportGraphicStyles

[![Direct](https://img.shields.io/badge/Direct%20Link-ImportGraphicStyles.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/style/ImportGraphicStyles.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ImportGraphicStyles.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- ダイアログの「スタイルを読み込み」ボタンで AI ファイルを指定し、そのファイル内のグラフィックスタイルを取り込む
- 取り込んだスタイル名でラジオボタンを自動生成し、選んだスタイルを選択オブジェクトへ適用する
- 指定したファイルは記憶され、次回以降は自動で参照（ボタンで別ファイルを選ぶまで保持）

### 処理の流れ

1. 対象オブジェクトを選択した状態でスクリプトを実行
2. 記憶しているファイルがあれば取り込み、取り込んだスタイル名でラジオを表示
3. 「スタイルを読み込み」ボタンで別ファイルを選ぶと、その場で取り込み直してラジオを更新
4. スタイルを選んで「適用」→ 選択オブジェクトへグラフィックスタイルを適用

### 注意点

- 取り込み時は「// _imported」レイヤーへ一時的に貼り付け、アセット登録後にレイヤーごと削除します（見た目は変わりません）
- 参照ファイルのパスは Folder.userData（swwwitch_ImportGraphicStyles.txt）に記憶し、スクリプト内には残しません
- ラジオは元ファイルのグラフィックスタイル名から自動生成されます（index 0 の既定スタイルは除外）

### 更新履歴

- v1.7.0 (20260701) : 参照ファイルをスクリプト固定から「スタイルを読み込み」ボタンでの選択式に変更（Folder.userData に記憶）。取り込んだスタイル名からラジオを自動生成する構成に変更。固定スタイル名（文字白抜き／枠のみ）を撤去
- v1.6.0 (20260701) : 検索・追加・候補リスト（ListBox／カテゴリ選択）を撤去。ラジオ（文字白抜き／枠のみ）で選んだグラフィックスタイルを、必要に応じて AI ファイルから取り込み、選択オブジェクトへ適用する構成に変更
- v1.5.0 (20260701) : ローカライズを構造化（ネスト LABELS ＋ ドット区切り L()）、全体を IIFE 化、パネル共通設定（setupPanel）を追加、変数名・関数名を整理、追加フローを関数分割、重複コード・不要な try を削減
- v1.4 (20250815) : 標準 ListBox（2 列ヘッダ）へ刷新、削除オプションを廃止、カテゴリラジオ＆検索ボタンを追加、貼り付け先を「// _imported」に統一、ドキュメント刷新
- v1.3 (20250815) : カテゴリ列（スタイル／ブラシ／シンボル／フォント）対応、追加時にカテゴリ選択（ドロップダウン）
- v1.2 (20250815) : CANDIDATES を外部 TSV から読み込み可能に
- v1.1 (20250815) : CANDIDATES に削除オプションを記録
- v1.0 (20250814) : 初期バージョン

### スクリプト情報

- バージョン: v1.7.0
