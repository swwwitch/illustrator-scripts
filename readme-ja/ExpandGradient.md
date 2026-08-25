# ExpandGradient

[![Direct](https://img.shields.io/badge/Direct%20Link-ExpandGradient.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/ExpandGradient.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ExpandGradient.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択オブジェクトに［オブジェクト］＞［分割・拡張］を実行し、グラデーションを指定したステップ数で分割します。

内部で一時的な .aia アクションを生成・ロード・再生し、実行後にアンロードして一時ファイルを削除します。

### 使い方

1. グラデーションを適用したオブジェクトを選択します。
2. スクリプトを実行します。
3. ステップ数と実行後の処理を指定して［OK］をクリックします。

### オプション

**ステップ数**

既定値 5、2 以上の整数。↑↓キーで±1、Shift+↑↓で10の倍数にスナップします。

**実行後の処理**

- なし: 分割・拡張のみ
- 単純に拡張: 分割・拡張のあとにアピアランスを分割
- クロップして拡張: Pathfinder のクロップを挟んでから拡張

### 注意点

- アクションのセット名／アクション名は定数で管理し、.aia 内の名前も同じ値から生成します。
- 後処理を簡略化した ExpandGradient-v2.jsx もあります。

### 更新履歴

- v1.1.0 (2026-05-25)
