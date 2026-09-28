# 配置画像の拡大・縮小率を表示して再スケール

[![Direct](https://img.shields.io/badge/Direct%20Link-ImageScaler.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/transform/ImageScaler.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ImageScaler.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択している配置画像（PlacedItem / RasterItem）の拡大・縮小率（%）を表示し、入力値で再スケールします。
- Rotation/Skew を含む変換行列から実スケールを算出し、入力フィールドでの↑↓キー操作による値変更にも対応します。

### 主な機能

- 実スケール（X/Y）を行列から算出
- ダイアログでスケール%を入力（単一選択時は現在値を初期表示）
- 入力値で相対倍率を計算して即時適用
- キーボード操作での数値調整：
  - ↑↓=次の整数へ（1.5→2）。入力欄の左のステップボタン（∧∨）のクリックも同じように効きます
  - Shift+↑↓=次の10の倍数へ（232→240）
  - Option(Alt)+↑↓=±0.1

### 処理の流れ

1. 選択オブジェクトを確認し、対象（配置画像／ラスタ画像）のみ抽出
2. 単一選択時は現在のスケール値を初期表示
3. ダイアログ表示、入力値の即時反映（app.redraw）
4. OKボタンで閉じる

### 更新履歴

- v1.0 (20250816) : 初期バージョン
- v1.1 (20250816) : キーボード操作による値変更機能追加
- v1.2 (20250816) : 入力値の即時適用（OKは閉じるのみ）、ローカライズ対応
- v1.3.2 (20260927) : ↑↓以外のキー入力で値が丸め直される不具合を修正。項目名にコロンを付け、ダイアログタイトルの余分な空白を削除
- v1.4.0 (20260927) : 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.4.1 (20260928) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.4.2 (20260928) : ボタン行を共通の部品で組むようにした

---

### スクリプト情報

- バージョン: v1.4.2
