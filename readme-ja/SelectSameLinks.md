# 同じリンクファイルを参照する配置画像を選択

[![Direct](https://img.shields.io/badge/Direct%20Link-SelectSameLinks.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/link/SelectSameLinks.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SelectSameLinks.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 現在選択中のリンク画像（PlacedItem）と同じリンクファイルを参照する
  PlacedItem をドキュメント内から抽出し、選択／削除を行う。
- Finds every PlacedItem in the active document that references the
  same linked file(s) as the current selection, then selects or deletes
  them according to the dialog options.

### ダイアログ / Dialog

- 判定方法 : 同じパス（絶対パス） / ファイル名一致（パス無視）
- 動作     : 同一リンクを選択 / リンク画像のみ削除 / クリップグループごと削除

### 更新履歴

- v1.1.7（2026-10-01）右側のボタンだけの行は、ダイアログの内側の幅（左右の余白を除く）が 200px 以内なら中央、それより広ければ右揃えに変更
- v1.1.6 (2026-09-30): 右側のボタンだけのボタン行を左右中央に並べるようにした
- v1.1.5 (2026-09-30): 文字ツールで文字を選択して実行するとエラーになる不具合を修正
- v1.1.4 (2026-09-29): ダイアログの不透明度を98%に変更
- v1.1.3 (2026-09-28): ボタン行を共通の部品で組むようにした
- v1.1.2 (2026-09-28): ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた

### スクリプト情報

- バージョン: v1.1.6
