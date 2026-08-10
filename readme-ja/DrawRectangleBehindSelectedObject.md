# DrawRectangleBehindSelectedObject

[![Direct](https://img.shields.io/badge/Direct%20Link-DrawRectangleBehindSelectedObject.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/DrawRectangleBehindSelectedObject.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DrawRectangleBehindSelectedObject.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択オブジェクトの外接バウンディングボックスを基準に、オフセットを加えた長方形を生成
- プレビュー機能で即時確認可能、作成した長方形は常に最背面に配置
- 不透明度はプレビューだけでなく、確定後に作成される長方形にも適用

更新日：2025-11-09

### 主な機能

- オフセット（現在の定規単位に追従）
- 角丸（ライブエフェクト適用、非展開）
- 塗り/線 カラー指定（K100 / ホワイト / HEX / CMYK）
- 対象：個別／グループとして
- プレビュー（専用レイヤー、ヒストリーに残らない）
- ダイアログ位置・不透明度・各パラメーターの記憶

### 処理の流れ

1. 対象オブジェクトのバウンディングボックスを計算
2. オフセット・角丸・塗り/線を適用した長方形を作成
3. プレビューは専用レイヤーに生成、確定時に本番描画
4. 「テキストとグループ化」オプションで元テキストとグループ化可能

### 更新履歴

- v1.0 (2025-08-22) : 初期バージョン
- v1.1 (2025-08-23) : プレビュー・カラー選択機能を追加
- v1.2 (2025-08-23) : 種別（塗り/線）、線幅指定、プリセット保存機能を追加
- v1.3 (2025-08-28) : ダイアログ位置・不透明度・各パラメーター記憶機能を追加
- v1.4 (2025-09-02) : ロジック調整
- v1.5 (2025-11-09) : 塗りロジックの見直し（HEX→CMYK変換、オーバープリント抑止、Normal固定）
- v1.6 (2025-11-09) : プレビュー安定化（debounce/cancelの整備、before/afterRender導入、bump互換、即時更新の不具合修正）

---

### スクリプト情報

- バージョン: v1.6
