# ImportGraphicStyles-v2

[![Direct](https://img.shields.io/badge/Direct%20Link-ImportGraphicStyles--v2.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/style/ImportGraphicStyles-v2.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ImportGraphicStyles-v2.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- ダイアログのラジオボタン（文字白抜き／枠のみ）でグラフィックスタイルを選択
- 選択したスタイルがドキュメントに無ければ、定義済みの AI ファイル（TARGET_FILE_PATH）から取り込む
- 選択中のオブジェクトへ、選んだグラフィックスタイルを適用する

### 処理の流れ

1. 対象オブジェクトを選択した状態でスクリプトを実行
2. ダイアログでスタイル（文字白抜き／枠のみ）を選択
3. スタイルが未登録なら対象 AI を開いてコピー→貼り付けで取り込み、一時オブジェクト・レイヤーを削除
4. 選択オブジェクトへグラフィックスタイルを適用

### 注意点

- 取り込み時は「// _imported」レイヤーへ一時的に貼り付け、アセット登録後にレイヤーごと削除します（見た目は変わりません）
- 読み込むファイルは TARGET_FILE_PATH、スタイル名は STYLE_NAME_WHITE_TEXT / STYLE_NAME_FRAME_ONLY で切り替えられます

### 更新履歴

- v1.6.0 (20260701) : 検索・追加・候補リスト（ListBox／カテゴリ選択）を撤去。ラジオ（文字白抜き／枠のみ）で選んだグラフィックスタイルを、必要に応じて AI ファイルから取り込み、選択オブジェクトへ適用する構成に変更
- v1.5.0 (20260701) : ローカライズを構造化（ネスト LABELS ＋ ドット区切り L()）、全体を IIFE 化、パネル共通設定（setupPanel）を追加、変数名・関数名を整理、追加フローを関数分割、重複コード・不要な try を削減
- v1.4 (20250815) : 標準 ListBox（2 列ヘッダ）へ刷新、削除オプションを廃止、カテゴリラジオ＆検索ボタンを追加、貼り付け先を「// _imported」に統一、ドキュメント刷新
- v1.3 (20250815) : カテゴリ列（スタイル／ブラシ／シンボル／フォント）対応、追加時にカテゴリ選択（ドロップダウン）
- v1.2 (20250815) : CANDIDATES を外部 TSV から読み込み可能に
- v1.1 (20250815) : CANDIDATES に削除オプションを記録
- v1.0 (20250814) : 初期バージョン

### スクリプト情報

- バージョン: v1.6.0
