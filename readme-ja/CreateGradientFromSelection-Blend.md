# 配置順の色からグラデーションを生成（ブレンド版）

[![Direct](https://img.shields.io/badge/Direct%20Link-CreateGradientFromSelection--Blend.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/CreateGradientFromSelection-Blend.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CreateGradientFromSelection-Blend.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択オブジェクトの塗り／線カラーを配置順（左→右、上→下）で抽出し、スウォッチグループに登録してグラデーションを自動生成します。ブレンド版です。

### 使い方

1. 色の元になるオブジェクトを選択します。
2. スクリプトを実行します。

### 注意点

- 値をセッション中だけ保持するため、専用エンジン（`#targetengine`）で動作します。Illustratorを再起動すると初期値に戻ります。
- 標準版は CreateGradientFromSelection.jsx です。

### 更新履歴

- v1.6
- v1.6.2 : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.6.3 : 一時アクションの読み込み・実行・解除を共通の処理にし、失敗してもアクションセットと一時ファイルが残らないようにした
- v1.6.3 : ボタン行を共通の部品で組むようにした
