# ビデオ定規・境界線などの表示を常駐パレットで切り替える

[![Direct](https://img.shields.io/badge/Direct%20Link-ViewTogglePalette.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/preference/ViewTogglePalette.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ViewTogglePalette.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

ビデオ定規・すべてのドキュメントでの定規・ガイド・スマートガイド・アートボード・境界線・カンバスカラー・バウンディングボックスの表示を、ボタン1つで切り替える常駐パレットです。

### 主な機能

| ボタン | 内容 |
| --- | --- |
| ビデオ定規 | ビデオ定規の表示／非表示 |
| すべてのドキュメントで表示 | 環境設定［すべてのドキュメントで定規を表示］（`useGlobalRulers`）のオン／オフ |
| ガイドを表示 | ガイドの表示／非表示 |
| ガイドをロック | ガイドのロック／ロック解除 |
| スマートガイド | スマートガイドのオン／オフ |
| アートボード | アートボードの表示／非表示 |
| 境界線 | エッジの表示／非表示 |
| カンバスカラー | カンバスカラーを「UIに合わせる」と「ホワイト」で切り替え |
| バウンディングボックス | バウンディングボックスの表示／非表示 |

### 使い方

スクリプトを実行するとパレットが開きます。ボタンを押すたびに表示が切り替わります。パレットは `Esc` キー（パレットがアクティブなとき）でも閉じられます。

すでにパレットが開いている状態でもう一度実行すると、新しく開き直さずに既存のパレットを前面に出します。

### 注意点

常駐パレットからは DOM を直接操作できないため、切り替えはすべて BridgeTalk 経由でメインエンジンに委譲しています。

「カンバスカラー」ボタンは `uiCanvasIsWhite` を書き換えたあとに `zoomout` → `zoomin` でキャンバスを描き直しています。

よく使う環境設定の切り替えは [AiQuickPrefsPalette](AiQuickPrefsPalette.md) が担当します。

### 更新履歴

- v1.0.0（2026-10-03）AiQuickPrefsPalette の表示パネルを別のスクリプトに分けた
- v1.0.1（2026-10-04）項目名のコロンを「 :」（半角スペース＋半角コロン）に変更（共通部品の更新）
- v1.0.2（2026-10-10）Illustrator を複数バージョン同時に起動していると、別のバージョンを操作してしまう不具合を修正
