# RegridObjects

[![Direct](https://img.shields.io/badge/Direct%20Link-RegridObjects.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/RegridObjects.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RegridObjects.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択中のオブジェクトが「だいたいグリッド状」に並んでいることを前提に、左右・上下の間隔で再配置します。
- 常時プレビュー対応。値は EditText で直接入力でき、↑↓ キー（Shift で ×10、Option で ×0.1）でも増減できます。
- 間隔の値は現在の定規単位（mm／pt／px など）で入力でき、内部では pt に換算して処理します（パネル名に単位を表示）。
- すでにグループになっているもの（クリップグループ含む）は、中身を分解せず「1 つのオブジェクト＝1 つの外接 bbox」として扱います。
- 実行後に自動でグループ化はしません（選択状態のまま）。
- ダイアログは日本語／英語の自動切り替えに対応（$.locale）。
- ［連動］で左右の値を上下に反映できます。
- ［レンガ状］：行ごとに左右位置を半ピッチずらして再配置します。
- ［ハニカム］：［レンガ状］と併用し、奇数行を（幅＋左右の値）の半分だけ横にずらし、行の高さだけ 0.75 倍にしてハニカム状にします（上下の値はそのまま反映）。
- ［強制グリッド］：位置の近さで列／行を推定せず、上→下の行ごとに左→右の順で (行, 列) を割り当てて再配置します。
- ［中央揃え］（［強制グリッド］のサブオプション）：各セル（列幅×行高さ）の天地左右中央にオブジェクトを整列します。
- ［行列入れ替え］はトグルです。ON で歯抜け（欠け）を許容しつつ行⇄列を転置し、OFF で転置前の状態に戻します。
- 1 行だけ→1 列、1 列だけ→1 行の転置にも対応します。

### 紹介記事（note）

https://note.com/dtp_tranist/n/n08861d0e40c3

### 更新履歴

- 2026-07-08: v1.6.0 中央揃え（強制グリッドのサブオプション）を追加、行列入れ替えをトグル化（OFFで転置前に戻る）、定規単位（mm/pt/px 等）での入力に対応（内部で pt 換算）、apply系関数の共通化・命名整理などコード整理
- 2025-10-31: v1.0 プレビューの常時 ON、連動（上下ディム）、↑↓ キーでの増減

### スクリプト情報

- バージョン: v1.6.0
