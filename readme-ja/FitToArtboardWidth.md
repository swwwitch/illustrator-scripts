# FitToArtboardWidth


[![Direct](https://img.shields.io/badge/Direct%20Link-FitToArtboardWidth.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/FitToArtboardWidth.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FitToArtboardWidth.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---


### 概要

選択したオブジェクト全体をひとまとまりとして、縦横比を保ったままアートボードの幅に合わせてリサイズし、アートボードの中央に配置します。既定の仕上がり幅はアートボード幅の90%です。

ダイアログはなく、実行するとそのまま処理します。[オブジェクトのリサイズ](SmartObjectResizer.md) の「アートボード → 幅」基準だけを、設定を選ばずに一発で実行できる形に切り出したものです。

### 主な機能

- 選択全体をひとまとまり（クラスタ）として、縦横比を保ったままリサイズ
- グループ化しないため、親階層や重ね順は変わりません
- 線幅・パターン・グラデーションも同じ倍率でスケール
- 基準にするアートボードは、選択と最も広く重なるものを自動判定
- リサイズ後にアートボードの中央へ配置
- 文字ツールで文字を選択している場合は、そのテキストオブジェクトを対象に切り替え

### 使い方

1. リサイズしたいオブジェクトを選択（複数可）
2. スクリプトを実行

### 設定

スクリプト冒頭の「ユーザー設定」ブロックで変更できます。

| 変数 | 既定値 | 説明 |
| --- | --- | --- |
| `WIDTH_PERCENT` | `90` | アートボード幅に対する仕上がり幅の割合（%）。`100` にするとアートボード幅ぴったり |
| `USE_PREVIEW_BOUNDS` | `true` | 計測に使う境界。`true` はプレビュー境界（線幅・効果込みの見た目の端）、`false` は幾何境界（パスの端） |
| `CENTER_VERTICALLY` | `true` | `false` にすると左右中央だけをそろえ、縦位置は動かしません |

### 処理の流れ

1. ドキュメントと選択状態を確認
2. 文字を選択している場合は、テキストオブジェクトを選択し直す
3. 選択全体を囲む矩形を求める
4. 基準にするアートボードを決める（選択と最も広く重なるもの、重ならなければ中心が最も近いもの）
5. アートボード幅 × `WIDTH_PERCENT` ÷ 100 を目標幅として倍率を計算
6. クラスタの左上を原点に、各オブジェクトのサイズと相対位置を同じ倍率で変形
7. アートボードの中央へ移動

### メモ

- 複数選択した場合は、オブジェクトどうしの間隔も同じ倍率で拡大縮小されます。**個々のオブジェクトをそれぞれアートボード幅に合わせたい場合は** [オブジェクトのリサイズ](SmartObjectResizer.md) **を使ってください。**
- `USE_PREVIEW_BOUNDS` が `true` のときは、線幅や効果を含めた見た目の端で計測するため、線幅も同じ倍率でスケールされます。
- アートボードが複数ある場合、アクティブなアートボードではなく、選択と最も広く重なるアートボードが基準になります。
- 高さ基準、裁ち落とし基準、片辺のみのリサイズが必要な場合は [オブジェクトのリサイズ](SmartObjectResizer.md) を使ってください。

### 更新履歴

- v1.0.0 (20260821) : 初版
