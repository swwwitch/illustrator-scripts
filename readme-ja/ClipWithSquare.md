# ClipWithSquare

[![Direct](https://img.shields.io/badge/Direct%20Link-ClipWithSquare.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/ClipWithSquare.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ClipWithSquare.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択された画像（配置画像/埋め込み画像）やクリッピングマスクグループ内の画像に対して、中心を基準とした最小の正方形パスを生成
- 生成した正方形で新しいクリッピンググループを作成

### 主な機能

- 配置画像、埋め込み画像、およびそれらを含むクリッピングマスクの処理
- ロックレイヤーやテンプレートレイヤー上の画像も対象（作業用レイヤーを作成して処理）

### 処理の流れ

1) 選択オブジェクトを走査
2) クリッピングマスクなら解除して画像のみ抽出
3) visibleBounds から最小正方形を作成
4) 正方形と画像を同一グループに入れてクリッピング化
5) 作成されたグループを選択状態に

### 更新履歴

- v1.0 (20231126) : 初期バージョン
- v1.1 (20250813) : クリッピング解除→再構築の安定化、コメント整理

### スクリプト情報

- バージョン: v1.1
