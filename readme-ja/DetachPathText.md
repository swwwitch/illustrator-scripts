# パス上文字をポイント文字に変換

[![Direct](https://img.shields.io/badge/Direct%20Link-DetachPathText.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/DetachPathText.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DetachPathText.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 更新日

20260928

### 概要

選択した「パス上文字（Path Text）」を、同じ文字内容・段落属性・文字属性をできるだけ維持したまま「ポイント文字」に変換します。
変換時に、「テキストの書式」およびパス設定を選択するダイアログが表示されます。
元のテキストパス形状は複製され、ダイアログの設定に応じて線属性（1pt黒／線なし）が適用されます。

### 使い方

1) パス上文字を選択
2) スクリプトを実行
3) ダイアログでテキストの書式およびパスのオプションを選択

### 注意点

- 文字ごとの属性（フォント/サイズ/色/トラッキング等）は、可能な範囲で復元します。
- 復元時に設定できない属性は try/catch でスキップします。

### スクリプト情報

- バージョン: v1.0.12

### 更新履歴

- v1.0.12 (2026-09-30) 右側のボタンだけのボタン行を左右中央に並べるようにした
- v1.0.11 (2026-09-30) 文字ツールで文字を選択して実行するとエラーになる不具合を修正
- v1.0.10 (2026-09-29) ダイアログの不透明度を98%に変更
- v1.0.9 (2026-09-28) ボタン行を共通の部品で組むようにした
- v1.0.8 (2026-09-28) ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
