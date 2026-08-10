# TextMergeToAreaBox-tab

[![Direct](https://img.shields.io/badge/Direct%20Link-TextMergeToAreaBox-tab.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/TextMergeToAreaBox-tab.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextMergeToAreaBox-tab.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 複数のテキストオブジェクトを1つのエリア内文字に連結します。
- 元のオブジェクトのサイズ・フォント・行送りを反映します。

### 主な機能

- 改行位置の調整（末尾が「。」以外、「.」「?」「!」の場合は連結）
- JUSTIFY 揃えの自動適用
- 元のオブジェクトの削除と置換処理

### 処理の流れ

1. 選択中のテキストオブジェクトを上から順にソート
2. 幅・高さ・行送りなどを取得
3. テキストを1つに連結し、エリア内文字を作成
4. 元のオブジェクトは削除

### 謝辞

倉田タカシさん（イラレで便利）
https://d-p.2-d.jp/ai-js/

### 更新履歴

- v1.0 (20250717) : 初期バージョン
- v1.1 (20250718) : 1行だけに対応、禁則を設定
- v1.2 (20250719) : 行末が英単語の場合の改行処理を追加

---

### スクリプト情報

- バージョン: v1.2
