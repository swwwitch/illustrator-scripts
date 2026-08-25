# FlattenOpacityPro

[![Direct](https://img.shields.io/badge/Direct%20Link-FlattenOpacityPro.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/FlattenOpacityPro.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FlattenOpacityPro.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択したオブジェクトの不透明度を、塗りのカラーそのものに焼き込んで不透明にします。

### 主な機能

- 親グループの不透明度も再帰的に合成
- 重なったオブジェクトを背面から合成して見た目の色を再現
- ブレンド方式を、リニアライトのRGB合成と元のカラースペースでの合成から切り替え可能（`USE_GAMMA_CORRECT_BLEND`）

### 使い方

1. 対象のオブジェクトを選択します。
2. スクリプトを実行します。

### 注意点

- 「同じ形状」とみなす許容差は `GEOM_TOL_PT`（位置・サイズ）と `AREA_TOL`（面積）で調整します。
- 元に戻せない変更を加えるため、実行前にファイルを複製しておくことを推奨します。

### 更新履歴

- v1.0
