# テキストの背面に図形を敷く

[![Direct](https://img.shields.io/badge/Direct%20Link-AddBackdrop.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/shape/AddBackdrop.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddBackdrop.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択したテキスト（またはオブジェクト）の背面に、見た目寸法に基づく図形を生成して配置します。

既存の背面図形があれば検出して置き換えます。

<img alt="背面に図形を敷くダイアログの外観" src="../png/ss-990-1102-144-20260923-190826.png" width="50%" />

### 主な機能

- 対象はポイント文字・エリア内文字。非テキスト選択時は選択範囲全体
- 形状は正円／スーパー楕円／長方形（キー操作 E / S / R）
- マージンは上下・左右を個別または連動で指定（既定は短辺の1/4）
- Undoベースのプレビューで、［OK］時に1ステップで確定

### 使い方

1. 背面に図形を敷きたいテキストまたはオブジェクトを選択します。
2. スクリプトを実行します。
3. 形状とマージンを指定して［OK］をクリックします。

### 紹介記事

https://note.com/dtp_tranist/n/na8af4a7016ad

### 更新履歴

- v1.6.4 (2026-09-23): UI文言を見直し（タイトル・パネル名・項目名のコロン・ツールチップ）、ボタンエリアを左右に分けて［透明グリッドを表示］ボタンを追加
- v1.6.3 (2026-09-23): 入力中の線幅とCMYK値が範囲外になっても補正されなかった不具合、角丸ONのとき前回の半径が復元されなかった不具合、ドキュメント未オープン時のメッセージを修正。内部処理を整理
- v1.6.1 (2026-03-26)
