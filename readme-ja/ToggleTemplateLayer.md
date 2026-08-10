# ToggleTemplateLayer

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

### スクリプト情報

- バージョン: v1.3
