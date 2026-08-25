# ExpandGradient-v2

[![Direct](https://img.shields.io/badge/Direct%20Link-ExpandGradient--v2.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/ExpandGradient-v2.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ExpandGradient-v2.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択オブジェクトに［オブジェクト］＞［分割・拡張］を実行し、グラデーションを指定したステップ数で分割します。ExpandGradient.jsx の簡易版です。

内部で一時的な .aia アクションを生成・ロード・再生し、実行後にアンロードして一時ファイルを削除します。

### 使い方

1. グラデーションを適用したオブジェクトを選択します。
2. スクリプトを実行します。
3. ステップ数を指定して［OK］をクリックします。

### オプション

**ステップ数**

既定値 5、2 以上の整数。↑↓キーで±1、Shift+↑↓で10の倍数にスナップします。

**実行後に拡張**

ONにすると、後処理として Pathfinder のクロップ（Live Pathfinder Crop）とアピアランスの分割（expandStyle）を順に実行します。

### 更新履歴

- v1.0.1 (2026-05-25)
