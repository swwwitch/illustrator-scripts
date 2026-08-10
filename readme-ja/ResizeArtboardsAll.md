# ResizeArtboardsAll

[![Direct](https://img.shields.io/badge/Direct%20Link-ResizeArtboardsAll.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/ResizeArtboardsAll.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ResizeArtboardsAll.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- ダイアログで指定した「幅」「高さ」に、アートボードをライブプレビューしながら変形します。
- 選択がない場合は、各アートボード内のオブジェクトを基準に、すべてのアートボードを個別に調整します。

### 主な機能

- ライブプレビュー（app.redraw のデバウンスで高速化）
- 対象のアートボード（作業のみ／すべて／指定（1始まりの範囲・カンマ列：例 1-3 / 1,3 / 2-4,7））
- 基準点の切替（左上／中央）
- 単位はドキュメントの定規単位に追従（px時は左上座標を整数にスナップ）
- ダイアログ位置・不透明度の記憶（セッション間で復元）
- 矢印キー操作：↑↓=±1、Shift+↑↓=10の倍数にスナップ、Option(Alt)+↑↓=±0.1（最終的に整数化）

### 処理の流れ

1. ダイアログで幅・高さを入力（必要なら対象・基準点を選択）
2. プレビューで即時にアートボードを更新
3. OKで確定、Cancelで元に戻す

### 注意点

- 「指定」は 1 始まりで解釈し、内部で 0 始まりに変換します。
- px/pt 表示は整数、小数単位（mm/cm/inch/pica）は小数2桁で表示します。

### 更新履歴

- v1.0 (20250829) : 初期バージョン

---

### スクリプト情報

- バージョン: v1.0
