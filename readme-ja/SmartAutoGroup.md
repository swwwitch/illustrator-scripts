# SmartAutoGroup

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartAutoGroup.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/group/SmartAutoGroup.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartAutoGroup.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択オブジェクトを自動的に「重なり」「垂直方向」「水平方向」「近接度」などの条件に応じてグループ化するIllustrator用スクリプトです。
- UI でモードとしきい値を指定し、再実行時には未グループオブジェクトを再確認できます。

### 主な機能

- 重なり・垂直・水平・近接度モード切替
- しきい値スライダーによる距離設定（重なりのみ除く）
- グループ化後の自動選択
- 未グループオブジェクトの再実行確認
- 日本語／英語インターフェース対応

### 処理の流れ

1. ダイアログでモードとしきい値を設定
2. DFS探索に基づきグループ単位を抽出
3. Illustrator標準 group コマンドで結合
4. グループ化されなかったオブジェクトがある場合に再実行を促す

### 更新履歴

- v1.0.0 (20250611) : 初期バージョン

---
