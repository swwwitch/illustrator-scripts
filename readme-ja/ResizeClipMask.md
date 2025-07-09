# マスクパスのサイズ変更

[![Direct](https://img.shields.io/badge/Direct%20Link-ResizeClipMask.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/ResizeClipMask.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- クリップグループ内のマスクパスを自動検出し、マージンを調整した新しいマスクに置き換えます。
- 複数クリップグループに一括適用可能で、長方形マスクのみ対応。

![](https://www.dtp-transit.jp/images/ss-442-260-72-20250710-044750.png)

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