# PDFをイラレで開いたときのバラバラ文字を、ひとつのエリア内文字に再構成するスクリプト

[![Direct](https://img.shields.io/badge/Direct%20Link-TextMergeToAreaBox.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/TextMergeToAreaBox.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要：

- 複数のテキストオブジェクトを1つのエリア内文字に連結します。
- 元のオブジェクトのサイズ・フォント・行送りを反映します。

### 主な機能：

- 複数のテキストオブジェクトを1つのエリア内文字に連結（元のテキストの幅を参照）
- 元のオブジェクトのサイズ・フォント・行送りを反映（行送りは見かけから計算。段落前後のアキには非対応）
- 改行位置の調整（末尾が「。」、「.」「?」「!」以外の場合には連結）
- 「両端揃え（最終行左揃え）」に設定
- 禁則「弱い」を適用しますが、［段落］パネルには表示されません。
- 1行だけの場合は、連結のみを行い、エリア内文字には変換しません。
- 行末が英単語、次の行頭が英単語の場合にスペースを挿入
- 行末が英単語とハイフン、次の行頭が英単語の場合にはハイフンを削除

### 処理の流れ：

1. 選択中のテキストオブジェクトを上から順にソート
2. 幅・高さ・行送りなどを取得
3. テキストを1つに連結し、エリア内文字を作成
4. 元のオブジェクトは削除

### 謝辞

倉田タカシさん（イラレで便利）
https://d-p.2-d.jp/ai-js/

### note

https://note.com/dtp_tranist/n/ne8d31278c266

### 更新履歴：

- v1.0 (20250717) : 初期バージョン
- v1.1 (20250718) : 1行だけに対応、禁則を設定
- v1.2 (20250719) : 行末が英単語の場合の改行処理を追加