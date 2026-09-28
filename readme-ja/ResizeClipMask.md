# マスクパスのサイズ変更

[![Direct](https://img.shields.io/badge/Direct%20Link-ResizeClipMask.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/ResizeClipMask.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ResizeClipMask.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- クリップグループ内のマスクパスを自動検出し、マージンを調整した新しいマスクに置き換えます。
- 複数クリップグループに一括適用可能で、長方形マスクのみ対応。

<img alt="" src="https://www.dtp-transit.jp/images/ss-442-260-72-20250713-082336.png" width="70%" />

### 主な機能

- マスクパス検出と選択
- ユーザー指定マージンのダイアログ入力
- 正負切替ボタンによる値反転
- 長方形判定とスキップ処理

### 処理の流れ

1. クリップグループを選択
2. ダイアログでマージンを指定
3. 長方形マスクを検出し、新しいマスクに置換
4. 元のマスクパスを削除

### 更新履歴

- v1.0.0 (20250710) : 初期バージョン
- v1.3.0 (20260927) : 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.3.1 (20260928) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.3.2 (20260928) : ボタン行を共通の部品で組むようにした。マスクの検出を共通の部品にした
- v1.3.3 (20260929) : ダイアログの不透明度を98%に変更