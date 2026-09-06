# 複数のオブジェクトを基準に「共通 > アピアランス」を実行

[![Direct](https://img.shields.io/badge/Direct%20Link-SelectSameAppearanceMulti.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/select/SelectSameAppearanceMulti.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SelectSameAppearanceMulti.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択したオブジェクトそれぞれを基準に［選択］メニューの「共通 > アピアランス」を実行し、見つかったオブジェクトをまとめて選択し直します。標準機能は基準にできるオブジェクトが1つだけですが、このスクリプトなら複数の基準を1回の実行で処理できます。

### 主な機能

- 選択中のオブジェクトを1つずつ基準にして「共通 > アピアランス」を実行
- 各回の結果を合算し、最後にすべてまとめて選択
- ダイアログなし。選択して実行するだけ
- 日本語／英語UI（警告メッセージのみ）

### 使い方

1. 基準にしたいオブジェクトを複数選択します。
2. スクリプトを実行します。
3. いずれかの基準とアピアランスが一致したオブジェクトが、すべて選択された状態になります。

たとえば「赤い線」と「青い塗り」を1つずつ選んで実行すると、赤い線のオブジェクトと青い塗りのオブジェクトが同時に選択されます。

### 注意点

- ドキュメントが開いていない、またはオブジェクトを選択していないときは、警告を表示して終了します。
- 基準にしたオブジェクト自身も結果に含まれます。
- 一致の判定は Illustrator の「共通 > アピアランス」そのままです。どこまでを同一と見なすかはアプリの挙動に従います。
- ロックまたは非表示のオブジェクトは選択できないため、結果に含まれません。
- 基準1つにつきメニューコマンドを1回実行します。選択数が多いと時間がかかります。
- 実行中は選択状態が一時的に切り替わります。処理が終わると結果の選択に置き換わり、元の選択には戻りません。

### 更新履歴

- v1.0.0 (20260906) : 初期バージョン
