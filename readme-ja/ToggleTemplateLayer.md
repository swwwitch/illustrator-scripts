# アクティブレイヤーのテンプレート属性を切り替える

[![Direct](https://img.shields.io/badge/Direct%20Link-ToggleTemplateLayer.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/layers/ToggleTemplateLayer.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ToggleTemplateLayer.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- アクティブレイヤーの「テンプレート」属性（ロック・印刷不可・画像を薄く表示）を ON / OFF する
- 小さなダイアログで ON（テンプレート化）／ OFF（解除）を選択
- ダイナミックアクションで実行
- 実行前にアクティブレイヤー名を取得し、リネームせず属性のみ適用
- ON はロックされたレイヤーには実行しない／ OFF はロック済み（テンプレート）でも実行
- 非表示レイヤーは ON / OFF とも対象外

### 更新履歴

- v1.0 (20240721) : 初期バージョン
- v1.1 (20260601) : 一時アクションの生成・実行を定型パターンに整理
- v1.2 (20260601) : アクティブレイヤー名を動的に取得して parameter-3 に注入
- v1.3 (20260601) : テンプレート OFF に対応し、ON/OFF を小ダイアログで選択
- v1.3.2 (20260928) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.3.3 (20260928) : 一時アクションの読み込み・実行・解除を共通の処理にし、失敗してもアクションセットと一時ファイルが残らないようにした
- v1.3.3 (20260928) : ボタン行を共通の部品で組むようにした
- v1.3.4 (20260929) : ダイアログの不透明度を98%に変更
- v1.3.5 (20260930) : 文字ツールで文字を選択して実行するとエラーになる不具合を修正
- v1.3.6 (20261001) : ボタン行の下に余白を加え、Illustrator 標準のダイアログに合わせた
- v1.3.7（2026-10-04）項目名のコロンを「 :」（半角スペース＋半角コロン）に変更（共通部品の更新）

### スクリプト情報

- バージョン: v1.3.7
