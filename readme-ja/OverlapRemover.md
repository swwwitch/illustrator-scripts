# OverlapRemover.jsx

[![Direct Link](https://img.shields.io/badge/Direct%20Link-OverlapRemover.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/path/OverlapRemover.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/OverlapRemover.md)

[![Back to home](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

## 概要

選択オブジェクトに対して、重なりをならすためのメニューコマンドを順に実行するスクリプトです。

オフセットパス → グループ化 → パスファインダー：合流 → アピアランスを分割、の順で処理します。オフセットパスで生成されたオブジェクトを元の選択と合わせて選び直すため、続くコマンドが意図した対象に効きます。

## 主な機能

- 選択オブジェクトのみを処理対象とする
- オフセットパスの実行前後を比較し、新しく作られたオブジェクトを検出
- 元の選択と新規オブジェクトを結合して後続のコマンドへ引き渡し
- ロック中・非表示のオブジェクトは選択対象から除外
- コマンドが失敗した場合はアラートを表示して中断
- 日本語／英語UI

## 使い方

1. 重なりをならしたいオブジェクトを選択します。
2. `OverlapRemover.jsx` を実行します。
3. オフセットパスのダイアログが開くので、オフセット量を入力して［OK］をクリックします。
4. 以降のグループ化、パスファインダー：合流、アピアランスを分割は自動で実行されます。

## 処理の流れ

| 順番 | 処理 | メニューコマンド |
| --- | --- | --- |
| 1 | オフセットパス | `OffsetPath v22` |
| 2 | グループ化 | `group` |
| 3 | パスファインダー：合流 | `Live Pathfinder Merge` |
| 4 | アピアランスを分割 | `expandStyle` |

## 注意事項

- オフセットパスはダイアログを開くため、値の指定はユーザー操作になります。
- ドキュメントが開かれていない場合、またはオブジェクトが選択されていない場合はメッセージを表示して終了します。
- ロック中・非表示のオブジェクトは選択対象に含まれません。
- Illustratorには取り消しをグループ化するAPIがないため、この処理の取り消しは複数ステップに分かれます。1回のUndoではすべて戻りません。

## 更新履歴

- v1.0.0 (20260301): 初期バージョン
- v1.0.1 (20260727): 「suspendHistoryによる単一Undoステップ化」という記述を削除（`suspendHistory` はPhotoshopのAPIで、Illustratorでは呼ばれないため）。あわせて基本情報・ローカライズのセクションを整備、日本語固定だったメッセージを日英対応に変更、選択判定の重複処理を `isSelectable` に集約
