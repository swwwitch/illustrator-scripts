# アートボード全体が正方形に近づくよう再配置

[![Direct](https://img.shields.io/badge/Direct%20Link-ArrangeArtboardsToSquareGrid.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/ArrangeArtboardsToSquareGrid.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArrangeArtboardsToSquareGrid.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

複数のアートボードを、全体の外形ができるだけ正方形（縦横比1:1）に近づく行列で再配置します。

### 主な機能

- 各アートボード内のアートワークも一緒に移動
- グリッドはカンバス中央に配置

### 使い方

1. 対象のドキュメントを開きます。
2. スクリプトを実行します。

### 注意点

- アートボード名から行列を決めたい場合は GridArrangeArtboards.jsx を使用してください。

### 更新履歴

- v1.0
- v1.1.0 (20260927) : 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.1.1 (20260928) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.1.2 (20260928) : 日本語表示で、アートボード数と推奨列数の表示のコロンを全角にそろえた。ボタン行を共通の部品で組むようにした
- v1.1.3 (20260929) : ダイアログの不透明度を98%に変更
- v1.1.4 (20260930) : 文字ツールで文字を選択して実行するとエラーになる不具合を修正
