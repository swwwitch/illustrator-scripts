# アートボードをKeynoteのスライドへPDF画像として挿入

[![Direct](https://img.shields.io/badge/Direct%20Link-ArtboardsToKeynoteAsPDF.applescript-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/AppleScript/ArtboardsToKeynoteAsPDF.applescript)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArtboardsToKeynoteAsPDF.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- Illustratorのアートボードを1枚ずつPDFで書き出し、Keynoteのスライドへ画像として挿入するAppleScript
- クリップボードとキー操作を使わないので速く、貼られるのはPDF画像（ベクトルのまま）

### 主な機能

- 対象のアートボードを「現在のアートボードのみ」「すべて」「範囲指定」から選択
- スケール（%）を指定して配置。100%でアートボードと同じ大きさ（1pt＝1ポイント）、初期値は全アートボードがスライドに収まる最大の値
- 「アートボード全体を書き出す」で、PDFの範囲をアートワークだけにするかアートボード全体にするかを選択
- 既存のドキュメントでは「末尾にスライドを追加して配置」か「既存スライドにオーバーレイ」を選択
- 画像はスライドの中央に配置

### 使い方

1. Illustratorでドキュメントを開きます。
2. スクリプトエディタなどからスクリプトを実行します。
3. ダイアログボックスで対象のアートボードと各項目を指定し、［挿入］をクリックします。

Keynoteでドキュメントが開いていなければ、新規ドキュメントを作ってスライド1から挿入します。

### オプション

| 項目 | 動作 |
| --- | --- |
| アートボード | 現在のアートボードのみ／すべて／範囲（〜）。初期値は現在のアートボードのみ |
| スケール | 書き出した範囲の大きさに対する倍率（%） |
| アートボード全体を書き出す | オンにすると、アートワークの外側の余白も含めてアートボード全体をPDFにします |
| 末尾にスライドを追加して配置 | アートボードごとに新しいスライドを末尾に追加します |
| 既存スライドにオーバーレイ | 開始スライドから順に、既存のスライドへ重ねて配置します。足りなくなったらスライドを末尾に追加します |
| 開始スライド | オーバーレイを始めるスライドの番号。初期値は表示中のスライド |

### 注意点

- 開始スライドの番号は、Keynoteのナビゲータに表示される番号（非表示スライドを数えない番号）です。オーバーレイでは非表示スライドを飛ばして配置します。
- オブジェクトのないアートボード、書き出しに失敗したアートボードは挿入せず、終了時に番号を表示します。
- 送り先のアプリケーションは「Keynote Creator Studio」です。
- 初回の実行時は、IllustratorとKeynoteを操作する許可（オートメーション）を求められます。

### 更新履歴

- v1.0.0 (2026-10-21) : 初版
- v1.0.1 (2026-10-22) : 非表示スライドがあると、指定したスライドより前に挿入される不具合を修正

### スクリプト情報

- バージョン: v1.0.1
