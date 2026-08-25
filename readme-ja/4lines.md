# 4lines

[![Direct](https://img.shields.io/badge/Direct%20Link-4lines.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/outline/4lines.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/4lines.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

アウトライン化した文字のパスを解析し、ディセンダーライン・ベースライン・ミーンライン・アセンダーラインの4本を推定して引きます。

### 使い方

1. アウトライン化した文字を選択します。
2. スクリプトを実行します。

### 注意点

- 解析にはアンカーポイントと水平セグメントの分布（ヒストグラム）を使います。十分なピークが見つからない場合は警告を表示して終了します。
- 複合パス（`o` や `A` のような穴のある字形）は、複合パス単位のバウンディングボックスとして扱います。
- 対象がアウトライン化されていない場合、PathItem が見つからないため処理できません。

### 更新履歴

- v1.0
