# QuickTransformPalette

[![Direct](https://img.shields.io/badge/Direct%20Link-QuickTransformPalette.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/transform/QuickTransformPalette.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/QuickTransformPalette.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択したオブジェクトの移動・複製と反転・回転を、アイコンのクリックで即時実行する常駐パレットです（プレビューや適用ボタンはありません）。

- 移動・複製: 上 / 左 / 右 / 下 の矢印アイコンをクリックでその方向へ移動。Option＋クリックで複製。十字の中央ボタンは移動せず、選択オブジェクトを同じ座標にその場で複製（複製後は複製側を選択）
- 反転・回転: アイコンボタンで左右反転／上下反転／90°回転（反時計回り・時計回り）を即時実行。Option＋クリックで複製してから変形。パネル最下部のスライダー（-180〜180°・15°刻み／Shift＋ドラッグで90°刻み）で任意角度に回転。スライダー左に現在の角度を数字で表示（離すと0°に戻る）。アイコン右の9軸（3×3）ウィジェットで基点を指定
- オプション（マージン／プレビュー境界。移動・複製と反転・回転の両方に適用）
	- マージン: 変形後に基準点の反対方向へ足す余白（単位は定規に追従）。反転・回転では9軸が中心のときは無視
	- プレビュー境界: 線や効果を含む見た目の境界でサイズ・基点を計算
- キーボード: Esc でパレットを閉じる／マージン欄は ↑↓ で増減（Shift＝±10・Option＝±0.1）。方向の移動・複製はアイコンのクリック操作のみ

DOM を触る処理（選択取得・移動・複製・反転・回転）はメインエンジンへ BridgeTalk で委譲します。

### 謝辞

kenさん
https://x.com/ken_rainy/status/1472505526768783361

### 紹介記事（note）

https://note.com/dtp_tranist/n/n277bd0865986

### スクリプト情報

- バージョン: v1.3.1
