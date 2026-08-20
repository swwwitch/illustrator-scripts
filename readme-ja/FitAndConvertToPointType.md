# FitAndConvertToPointType


[![Direct](https://img.shields.io/badge/Direct%20Link-FitAndConvertToPointType.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/FitAndConvertToPointType.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FitAndConvertToPointType.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---


### 概要

選択したエリア内文字をポイント文字に変換するスクリプト。テキストがあふれているものだけ自動サイズ調整をONにしてあふれを解消してから変換するため、あふれた文字が失われません。ダイアログで［強制改行を削除］を選ぶと、変換後に残る強制改行を取り除いて行をつなげます。

### 主な機能

- 選択の中からエリア内文字だけを対象にする（ロック・非表示のもの、ロック／非表示レイヤー上のものは除外）
- あふれているフレームだけ自動サイズ調整をONにして、あふれを解消
- ポイント文字に変換
- ［強制改行を削除］で、変換後に残る強制改行（ソフトリターン）を削除して行をつなげる（段落改行は残る）
- 変換後のポイント文字を選択し直す
- 変換できなかったものがある場合は件数を表示

### 処理の流れ

1. ドキュメント・選択内容・変換APIの対応状況を確認
2. ダイアログを表示（［強制改行を削除］）
3. 自動サイズ調整用のアクションセットを一時的に読み込む
4. あふれているフレームだけ選択し、自動サイズ調整を実行
5. ポイント文字に変換（［強制改行を削除］がONなら、変換後に強制改行を削除）
6. アクションセットを破棄し、変換後のポイント文字を選択

### メモ

- 自動サイズ調整はスクリプト（DOM）から設定できないため、一時的なアクションセットを読み込んで実行しています。アクションセット名にはスクリプト名を冠して（`FitAndConvertToPointType_AutoSize`）既存のアクションと衝突しないようにし、処理の成否にかかわらず終了時に破棄します。
- あふれていないフレームには自動サイズ調整をかけません。テキストの配置が［中央］［下］のときに文字が動いてしまうためです。
- ［強制改行を削除］で削除するのは強制改行（ソフトリターン：U+0003 / U+000A / U+2028）だけで、段落改行（U+000D）は残ります。文字ごとの書式を保つため、`contents` の一括置換ではなく1文字ずつ削除しています。
- ［強制改行を削除］の初期状態はOFFです。常にONで始めたい場合は、スクリプト冒頭の `DEFAULT_REMOVE_LINE_BREAKS` を `true` に変更してください。
- 連結（スレッド）されたテキストなど、Illustratorがポイント文字に変換できないものは対象外です。その場合は「◯件を変換しました。◯件は変換できませんでした」と表示されます。
- ポイント文字への変換APIを持たない古いIllustratorでは、その旨を表示して何も変更しません。
- 変換APIは変換後のオブジェクトを返さないため、対象に一時的な目印の名前を付けて回収しています。処理後は元の名前へ戻します。

### 更新履歴

- v1.0.0 (20260820) : 初期バージョン（ダイアログで［強制改行を削除］を選択可）
