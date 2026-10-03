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
- v1.3.4 (20260930) : 文字ツールで文字を選択して実行するとエラーになる不具合を修正
- v1.3.5 (20260930) : 右側のボタンだけのボタン行を左右中央に並べるようにした
- v1.3.6（2026-10-01）右側のボタンだけの行は、ダイアログの内側の幅（左右の余白を除く）が 200px 以内なら中央、それより広ければ右揃えに変更
- v1.3.7（2026-10-01）ウィンドウ・パネルの余白と間隔を共通部品（UIレイアウト）にそろえた
- v1.3.8（2026-10-01）ボタン行の下に余白を加え、Illustrator 標準のダイアログに合わせた
- v1.3.9（2026-10-03）数値欄の中に単位を表示し、項目名の単位表記を削除