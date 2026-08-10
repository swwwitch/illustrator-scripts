# RenameAssets

[![Direct](https://img.shields.io/badge/Direct%20Link-RenameAssets.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/RenameAssets.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RenameAssets.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- グラフィックスタイル／ブラシ／スウォッチ／シンボルの「名前」を、ダイアログで指定した「検索→置換」で一括変更します。
- プレビュー（ボタン押下時のみ）で事前に変更結果を確認できます。キャンセルで元に戻せます。

### 主な機能

- 検索・置換、正規表現対応、大文字小文字の無視
- 対象（スタイル／ブラシ／スウォッチ／シンボル）の切替（ラジオボタン）
- グラフィックスタイルに限り、対象スタイルの**複数選択フィルタ**
- 重複名の自動回避（ (2), (3), … を付与）
- 角括弧で囲まれた既定名のスキップ
- ダイアログの透明度と表示位置の調整
- 日本語／英語のローカライズ

### 処理の流れ

1. ダイアログで対象コレクションと検索／置換文字列、オプション（正規表現・大文字小文字）を設定
2. （スタイルの場合）必要に応じて対象スタイルを複数選択
3. ［プレビュー］を押してパネル上の名称に一時反映（スナップショットから安全に復元可能）
4. 問題なければ［OK］で確定、［キャンセル］で元に戻す

### 注意点

- ライブ（入力中）プレビューは行いません。プレビューはボタン押下時のみ実行されます。

### 更新履歴

- v1.0 (20250820) : 初期バージョン

### スクリプト情報

- バージョン: v1.0
