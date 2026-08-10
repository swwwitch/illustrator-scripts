# OutlineTextRestore

[![Direct](https://img.shields.io/badge/Direct%20Link-OutlineTextRestore.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/outline/OutlineTextRestore.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/OutlineTextRestore.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- アウトライン化されたテキストを、メモ情報（note）をもとに元のテキストとして復元するスクリプトです。
- 選択しているパス／グループのみを対象とし、復元されたテキストは「restored_text」レイヤーに統合して作成されます。
- 元のアウトライン（選択オブジェクトのみ）は、新規作成した退避レイヤーに移動され、そのレイヤーを「outlined_text」として最背面に配置し、テンプレートレイヤーに設定します。

### 更新日

20260111

### 主な機能

- PathItem または GroupItem の note 情報からテキスト内容、フォント情報、座標などを抽出
- 新しいテキストフレームを元の位置に再生成し、元のアウトラインを非表示または移動
- メモが存在しない・情報が不正な場合には警告を表示
- 復元テキストは毎回「新規レイヤー」に作成／元アウトラインは毎回「新規退避レイヤー」に移動して outlined_text（最背面・テンプレート化）

### 処理の流れ

1. 対象オブジェクトを選択（アウトライン化された Path または Group）
2. メモ情報を解析し、属性データを抽出
3. 復元テキストを「新規レイヤー」に再構築
4. 元オブジェクトを「outlined_text」レイヤーへ移動（このレイヤーは常に最背面・テンプレート化）

### 更新履歴

- v1.0 (20240811) : 初期バージョン
- v1.1 (20250721) : 微調整
- v1.2 (20260111) : ローカライズ（アラート文言の英語対応）
- v1.3 (20260111) : 退避レイヤー名を outlined_text に変更（旧名 outlined-text から自動移行）

### スクリプト情報

- バージョン: v1.3
