# 複数のテキストを1つのエリア内文字に連結

[![Direct](https://img.shields.io/badge/Direct%20Link-TextMergeToAreaBox--tab.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/TextMergeToAreaBox-tab.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextMergeToAreaBox-tab.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 分割されたテキストオブジェクトを行単位にまとめ、1つのエリア内文字に連結します。
- 同じ行の断片はタブで区切り、行ごとに改行します。表組みの再構成に使います。
- 元のテキストのフォント・サイズ・行送りを引き継ぎます。

### 主な機能

- Y座標が近いオブジェクトを同じ行としてまとめ、行内は左から右の順にタブ区切りで連結
- 行ごとに改行（行の並びをそのまま保つ）
- 断片の中の改行は「〓」に置き換え（行の区切りと混ざらないように）
- 禁則「弱い禁則 v2」（使えないバージョンでは「弱い禁則」）と、両端揃え（最終行左揃え）をテキスト全体に適用
- 最終行があふれる場合は枠を下に伸ばす
- 1行だけのときはエリア内文字にせず、左揃えで出力
- 元のオブジェクトの削除と置換処理

### 処理の流れ

1. 選択中のテキストオブジェクトを上から順に並べ、行ごとに分ける
2. 行ごとに左から右へタブ区切りで連結
3. 全体の外接矩形（幅は1文字分縮める）からエリア内文字を作成し、フォント・サイズ・行送りを引き継ぐ
4. 元のオブジェクトは削除

### 謝辞

倉田タカシさん（イラレで便利）
https://d-p.2-d.jp/ai-js/

### 更新履歴

- v1.0 (2025-07-18) : 初期バージョン
- v1.1 (2025-07-19) : 1行だけに対応、禁則を設定
- v1.2 (2025-07-20) : 行末が英単語の場合の改行処理を追加
- v1.2.2 (2026-09-21) : 行は単純に改行でつなぐよう修正（文末の後の空行・行頭の半角スペースを解消）、禁則を「弱い禁則 v2」に変更（使えないバージョンでは「弱い禁則」）、最終行があふれる場合に枠を下に伸ばすよう修正、幅が狭い選択で作成に失敗しないよう修正、バウンディングボックスを選択状態に依存しないよう修正、1行のときも横組みに統一、ドキュメント未オープン時のガード追加、コードを整理

---

### スクリプト情報

- バージョン: v1.2.2
