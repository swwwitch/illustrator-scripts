# AddOutlineOffsetPath

[![Direct](https://img.shields.io/badge/Direct%20Link-AddOutlineOffsetPath.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/path/AddOutlineOffsetPath.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddOutlineOffsetPath.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 更新日：2025-08-13
- 選択オブジェクトを複製 → 背面配置 → オフセットパス（Live Effect）→ アウトライン → 合体 → 拡張
- 元オブジェクトと結果をグループ化し、Subtract を実行して白で塗りつぶす

### 主な機能

- 複数選択対応
- 単位対応（pt, mm, in, cm など）
- 角の形状（マイター、ラウンド、ベベル）設定可能
- オフセット値のダイアログ入力
- ダイアログの位置調整と透明度設定
- Shift/Option キーによる数値入力の増減制御

### 処理の流れ

1) 選択オブジェクトを複製し、背面へ移動
2) オフセットパス（Live Effect）を適用
3) アウトライン化し、合体（Unite）後に拡張（Expand）
4) 元オブジェクトと結果をグループ化し、Subtract を実行
5) 結果を白で塗りつぶす

### クレジット

このスクリプトの一部は、以下のスクリプトを参考にして開発しました。
Outline.jsx (illustrator-outline-script) 作者: Oğuzhan Yıldırım @oguzhanyildirim01
https://github.com/oguzhanyildirim01/illustrator-outline-script/blob/main/Outline.jsx

### 更新履歴

- v1.0.0 (2025-08-13) : 初期バージョン
- v1.1.0 (2025-08-13) : オフセット値の自動計算

### スクリプト情報

- バージョン: v1.1.0
