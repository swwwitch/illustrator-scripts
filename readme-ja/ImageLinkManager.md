# ImageLinkManager

[![Direct](https://img.shields.io/badge/Direct%20Link-ImageLinkManager.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/link/ImageLinkManager.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ImageLinkManager.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

配置画像（PlacedItem）に対する「埋め込み」「埋め込み解除」「リセット」「ケイ線」「リンク」を、1つのダイアログでまとめて扱います。

ダイアログ上部のモードで処理を切り替え、対応するパネルだけが有効になります。

### 主な機能

- 埋め込み: 選択またはドキュメント内すべての PlacedItem を `embed()`（PSD は埋め込み後に ungroup を試行）
- 埋め込み解除: 埋め込み画像をリンク画像に戻す
- リセット: 配置画像の変形をリセット
- ケイ線: 配置画像にケイ線を追加
- リンク: リンク先の付け替え
- モード切替ショートカット: 埋め込み **E** / 解除 **U** / リセット **R** / ケイ線 **S** / リンク **L**

### 使い方

1. 対象の配置画像を選択します（ドキュメント全体を対象にできるモードもあります）。
2. スクリプトを実行します。
3. 上部でモードを選び、パネルの設定を指定して実行します。

### 注意点

- 一覧表示を伴う本格的な管理には LinkedImageManager.jsx を使用してください。

### 更新履歴

- v1.2 (2025-12-21)
