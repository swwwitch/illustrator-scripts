# ドキュメント内の日付を検索して置換

[![Direct](https://img.shields.io/badge/Direct%20Link-DateFindReplace.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/DateFindReplace.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DateFindReplace.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

ドキュメント内のテキストフレームから日付を検索し、選択した項目だけを置換します。

### 使い方

1. （必要に応じて）検索対象を絞るためにオブジェクトを選択します。
2. スクリプトを実行します。
3. 見つかった日付から置換するものを選び、新しい日付を指定して実行します。

### 注意点

- オブジェクトが選択されている場合は、その選択範囲（グループ内を含む）のテキストフレームだけを検索対象にします。

### 更新履歴

- v1.1.7（2026-10-01）ウィンドウ・パネルの余白と間隔を共通部品（UIレイアウト）にそろえた
- v1.1.6（2026-10-01）右側のボタンだけの行は、ダイアログの内側の幅（左右の余白を除く）が 200px 以内なら中央、それより広ければ右揃えに変更
- v1.1.5 (20260930) : 右側のボタンだけのボタン行を左右中央に並べるようにした
- v1.1.4 (20260930) : 文字ツールで文字を選択して実行するとエラーになる不具合を修正
- v1.1.3 (20260929) : ダイアログの不透明度を98%に変更
- v1.1.2 (20260928) : ボタン行を共通の部品で組むようにした
- v1.1.1 (20260928) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.1.0 (20260927) : 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.0.4 (20260927) : 英語UIに対応。Option＋クリックの一括切替と、Escで閉じたときのプレビューの巻き戻しを修正。↑↓キー以外の入力で値が丸め直されないように修正
- v1.0.2
