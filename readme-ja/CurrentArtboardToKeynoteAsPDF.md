# 作業中のアートボードをKeynoteの表示中のスライドへ挿入

[![Direct](https://img.shields.io/badge/Direct%20Link-CurrentArtboardToKeynoteAsPDF.applescript-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/AppleScript/CurrentArtboardToKeynoteAsPDF.applescript)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CurrentArtboardToKeynoteAsPDF.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 作業中のアートボードをPDFで書き出し、Keynoteの表示中のスライドへ画像として重ねるAppleScript
- ダイアログボックスを出さない版。[ArtboardsToKeynoteAsPDF](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArtboardsToKeynoteAsPDF.md) の既定の動作を、スライドに収まる大きさで行います

### 主な機能

- アートボード全体の範囲でPDFを書き出し、スライドに収まる最大の大きさで中央に配置
- 挿入後、Keynoteで挿入先のスライドを表示

### 使い方

1. Keynoteで挿入先のスライドを表示しておきます。
2. Illustratorで挿入したいアートボードを作業中にします。
3. スクリプトを実行します。

Keynoteでドキュメントが開いていなければ、新規ドキュメントを作って挿入します。

### 注意点

- オブジェクトのないアートボードは挿入せず、警告を表示します。
- 送り先のアプリケーションは「Keynote Creator Studio」です。
- 初回の実行時は、IllustratorとKeynoteを操作する許可（オートメーション）を求められます。

### 更新履歴

- v1.0.0 (2026-10-21) : 初版
- v1.0.1 (2026-10-22) : 非表示スライドがあると、表示中より前のスライドに挿入される不具合を修正

### スクリプト情報

- バージョン: v1.0.1
