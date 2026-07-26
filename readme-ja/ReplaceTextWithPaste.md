# ReplaceTextWithPaste.jsx

[![Direct Link](https://img.shields.io/badge/Direct%20Link-ReplaceTextWithPaste.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/ReplaceTextWithPaste.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReplaceTextWithPaste.md)

[![Back to home](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

## 概要

クリップボードにあるテキストで、選択中のテキストフレームの内容をまとめて置き換えるスクリプトです。

選択がない場合は、コピー元と同じ位置に新しいテキストフレームを作成します。グループを選択した場合は、その中のテキストフレームをすべて対象にします。

## 主な機能

- 選択したテキストフレームの内容を一括置換
- 選択なしのときは、コピー元と同じ位置に新規テキストフレームを作成
- グループ内のテキストフレームを再帰的に処理
- ポイント文字・エリア内文字・パス上文字に対応
- 処理の前後で選択状態を保持
- 日本語／英語UI

## 使い方

1. 置き換えたいテキストを、テキストエディタなどでコピーします。
2. 置き換え先のテキストフレームを選択します（選択しない場合は新規作成されます）。
3. `ReplaceTextWithPaste.jsx` を実行します。

## 処理の流れ

1. 現在の選択を退避します。
2. 一度通常のペーストを実行し、貼り付けられたテキストフレームから内容と座標を取得します。
3. 貼り付けたオブジェクトを削除し、退避した選択を復元します。
4. 選択があれば各テキストフレームの内容を置き換え、なければ新規テキストフレームを作成します。

## 対象

| 区分 | 対象 |
| --- | --- |
| 対象 | テキストフレーム、グループ内のテキストフレーム |
| 非対象 | 画像、図形、ロックされたオブジェクト |

## 注意事項

- 内部で一度通常のペーストを実行します。クリップボードにテキストが入っていない場合は何も起こりません。
- 文字揃えや文字スタイルは元のテキストフレームの設定が保持されます。置き換わるのは内容だけです。
- Illustratorには取り消しをグループ化するAPIがないため、この処理の取り消しは複数ステップに分かれます。
- 原案：Gorolib Design

## 更新履歴

- v1.0.0 (20260727): バージョン変数を導入。効果のなかった `doc.undoGroup` への代入を削除し、選択復元とクリップボード取得の重複処理を整理。ドキュメント未オープン時のメッセージを追加
