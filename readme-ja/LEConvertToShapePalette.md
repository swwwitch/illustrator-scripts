# ［形状に変換］のライブエフェクトを適用するパレット

[![Direct](https://img.shields.io/badge/Direct%20Link-LEConvertToShapePalette.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/single-function/LEConvertToShapePalette.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/LEConvertToShapePalette.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択オブジェクトに［形状に変換］のライブエフェクトを適用する常駐パレット。パレットで「長方形／楕円」と「値を指定（Absolute）／値を追加（Relative）」、幅・高さ（pt）を設定すると、選択にライブプレビューが反映される。DOM 操作は BridgeTalk でメインエンジンへ委譲する。

### 更新履歴

- v1.1.6（2026-10-10）Illustrator を複数バージョン同時に起動していると、別のバージョンを操作してしまう不具合を修正
- v1.1.5（2026-10-04）項目名のコロンを「 :」（半角スペース＋半角コロン）に変更（共通部品の更新）
- v1.1.4（2026-10-03）常駐エンジンを共通エンジン `SwwwitchPalettes` に変更し、［常駐パレットをまとめて閉じる］で閉じられるようにした
- v1.1.3（2026-10-01）ボタン行の下に余白を加え、Illustrator 標準のダイアログに合わせた
- v1.1.2（2026-10-01）［適用］ボタンをパネル内からパレット下段の標準のボタン行へ移した。ウィンドウ・パネルの余白と間隔を共通部品（UIレイアウト）にそろえた
- v1.1.1（2026-09-28）英語 UI の項目名のコロンの後ろの空白をなくした（ローカライズ処理の共通化）
- v1.1.0（2026-09-27）数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.0.1（2026-09-26）ファイル名を `LEConvertToShape.jsx` から `LEConvertToShapePalette.jsx` に変更。
- v1.0.0: 常駐パレット化。形状（長方形／楕円）とサイズモード（値を指定／値を追加）＋幅・高さを設定し、ライブプレビュー付きで［形状に変換］を適用 ／ Persistent palette. Choose shape (rectangle/ellipse) and size mode (absolute/relative) plus width/height, apply "Convert to Shape" with a live preview.

### スクリプト情報

- バージョン: v1.1.6
