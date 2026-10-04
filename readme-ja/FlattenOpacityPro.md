# 不透明度を塗りのカラーに焼き込む

[![Direct](https://img.shields.io/badge/Direct%20Link-FlattenOpacityPro.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/FlattenOpacityPro.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FlattenOpacityPro.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択したオブジェクトの不透明度を、塗りのカラーそのものに焼き込んで不透明にします。
グラデーションや画像など焼き込めないものを含むときは、標準の［透明部分を分割・統合］で処理します。

### 主な機能

- 親グループの不透明度も再帰的に合成
- 重なったオブジェクトを背面から合成して見た目の色を再現
- 焼き込めないもの（グラデーション・パターン・グレー・特色の色、テキスト・画像・シンボルなど、描画モードが通常以外）を含むときは、選択全体に標準の［透明部分を分割・統合］を適用
- ブレンド方式を、リニアライトのRGB合成と元のカラースペースでの合成から切り替え可能（`USE_GAMMA_CORRECT_BLEND`）

### 使い方

1. 対象のオブジェクトを選択します。
2. スクリプトを実行します。

### 注意点

- 「同じ形状」とみなす許容差は `GEOM_TOL_PT`（位置・サイズ）と `AREA_TOL`（面積）で調整します。
- 効果（ドロップシャドウなど）は判定できないため、焼き込みの対象になります。
- 元に戻せない変更を加えるため、実行前にファイルを複製しておくことを推奨します。

### 更新履歴

- v1.0
- v1.0.2（2026-09-27）アラートを英語表示にも対応
- v1.1.0（2026-09-29）FlattenTransparency.jsx を統合。焼き込めないものを含むときは標準の［透明部分を分割・統合］で処理するようにした
- v1.1.1（2026-10-04）項目名のコロンを「 :」（半角スペース＋半角コロン）に変更（共通部品の更新）
