# 行送りと配置を上下方向に調整するパレット

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartDistributorPalette.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/SmartDistributorPalette.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartDistributorPalette.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

DistributeDownFromTop.jsx / DistributeUpFromTop.jsx を統合した常駐パレットです。

十字ボタン（↑ / ← 0 → / ↓）を押すたびに、その時点の選択へ1ステップぶん適用します。

### 使い方

1. 対象のオブジェクトまたはテキストを選択します。
2. スクリプトを実行してパレットを開きます。
3. 十字ボタンを押して調整します。

### 注意点

- 実際のドキュメント操作は BridgeTalk でメインエンジンへ送って実行するため、1クリック＝取り消し1回になります。
- 常駐パレット型のため、スクリプトを修正したあとは**パレットを閉じてから再実行**してください。

### 更新履歴

- v1.1.6（2026-10-04）項目名のコロンを「 :」（半角スペース＋半角コロン）に変更（共通部品の更新）
- v1.1.6（2026-10-04）項目名のコロンを「 :」（半角スペース＋半角コロン）に変更（共通部品の更新）
- v1.1.5（2026-10-03）常駐エンジンを共通エンジン `SwwwitchPalettes` に変更し、［常駐パレットをまとめて閉じる］で閉じられるようにした
- v1.1.4（2026-10-01）ウィンドウ・パネルの余白と間隔を共通部品（UIレイアウト）にそろえた
- v1.1.3（2026-09-29）Esc で閉じるようにした（入力中も効く）
- v1.1.2（2026-09-29）メインエンジンへ送るワーカーのソースを終わりの目印で切り詰めるようにした（前にコードを足すと送信本文が壊れて動かなくなるのを防ぐ）
- v1.1.1（2026-09-28）設定の保存を共通の部品にした（保存先: `~/Library/Application Support/illustrator-scripts/SmartDistributorPalette.json`。旧版の設定は最初の1回だけ読み継ぐ）
- v1.1.0（2026-09-27）移動距離「カスタム」の数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.0.3（2026-09-26）ファイル名を `SmartDistributor.jsx` から `SmartDistributorPalette.jsx` に変更。
- v1.0.2 (2026-09-19) 綴り違いの重複 `SmartDistributer.jsx` を統合
- v1.0.1
