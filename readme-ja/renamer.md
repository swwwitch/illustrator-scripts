# アートボード・レイヤー・シンボル名を検索置換とナンバリングで変更

[![Direct](https://img.shields.io/badge/Direct%20Link-renamer.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/renamer.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/renamer.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- アートボード／レイヤー／シンボル／グラフィックスタイルの名前を、検索置換とナンバリングでまとめて変更します。
- ダイアログで対象を切り替え、変更結果を一覧で確認しながら調整できます。

### 主な機能

- 検索・置換（正規表現に対応）
- 接頭辞・接尾辞の付与
- ナンバリング（区切り文字・開始番号を指定）
- 並び替え（元の順／名前昇順／名前降順／変更あり優先）
- 一覧内での上下移動（↑↑ ↑ ↓ ↓↓）

### 使い方

1. スクリプトを実行する
2. 対象（アートボード／レイヤー／シンボル／グラフィックスタイル）を選ぶ
3. 検索置換・接頭辞／接尾辞・ナンバリングを設定する
4. 一覧で結果を確認して［OK］

### 注意点

- ドキュメントが開かれていない場合はアラートを表示して終了します。

### 更新履歴

- v1.1.4 (20260930) : 文字ツールで文字を選択して実行するとエラーになる不具合を修正
- v1.1.3 (20260929) : ダイアログの不透明度を98%に変更
- v1.1.2 (20260928) : ボタン行を共通の部品で組むようにした
- v1.1.2 (20260928) : 英語表示で、項目名のコロンの後ろの余分な空白と、件数の後ろに出ていた「件」を削除
- v1.1.1 (20260928) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.1.0 (20260927) : 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
