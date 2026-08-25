# LogoGridMaker

[![Direct](https://img.shields.io/badge/Direct%20Link-LogoGridMaker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/LogoGridMaker.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/LogoGridMaker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択オブジェクトの visibleBounds を基準に、ロゴ用の補助線およびクリアスペースを生成します。

「文字形状からグリッド構造を抽出する」ことを目的としたスクリプトです。

選択範囲の境界に一致する線は、太く強調して表示します。

### 主な機能

- 横線: 横伸張率、上下に線を追加、左方向に延長。ラインの決め方は なし／自動判定／水平セグメント／均等分割
- 縦線: 縦伸張率、左右に線を追加、縦分割、上方向に延長。垂直エレメント／斜線エレメントの抽出に対応
- 共通設定とクリアスペース: 分割数を基準にユニットを定義してクリアスペースを生成。作成レイヤー名・線幅・ガイド化・グループ化を指定
- プリセット: 1x1 / auto / element / left / up-3 / clear space

### 使い方

1. ロゴなど、基準にするオブジェクトを選択します。
2. スクリプトを実行します。
3. プリセットまたは個別の設定を指定して実行します。

### 注意点

- クリアスペースをONにすると、横線・縦線はパネルごと無効になります。このとき線幅・ガイド化は無効、グループ化はON固定です。
- 現在のUI状態を、JSON（配列貼り付け形式）として書き出せます。

### 更新履歴

- v1.4.1 (2026-04-10)
