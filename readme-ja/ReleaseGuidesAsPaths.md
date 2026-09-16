# 「_guide」レイヤーのガイドをパスに戻して移動

[![Direct](https://img.shields.io/badge/Direct%20Link-ReleaseGuidesAsPaths.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/guide/ReleaseGuidesAsPaths.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReleaseGuidesAsPaths.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

「_guide」レイヤー内のガイドを通常のパスに戻し、「ReleasedGuides」レイヤーへ移動します。

- 「_guide」レイヤーのロックを解除します
- レイヤー内のアイテムのガイドを解除し、塗りなし・線K100・1ptにします
- 「ReleasedGuides」レイヤーへ移動します（なければ作成します。以前の「UnlockedGuides」レイヤーがあれば、そちらを使います）
- 移動後、「_guide」レイヤーを再ロックします

### 更新履歴

- v1.0.1 (20260916) : スクリプト名を unlockGuideLayerAndClearGuides から ReleaseGuidesAsPaths に変更。線がK100にならない不具合を修正。ドキュメントや「_guide」レイヤーがないときにアラートを表示。移動先レイヤー名を「ReleasedGuides」に変更（既存の「UnlockedGuides」レイヤーはそのまま使用）
- v1.0 (20250716) : 初期バージョン

### スクリプト情報

- バージョン: v1.0.1
