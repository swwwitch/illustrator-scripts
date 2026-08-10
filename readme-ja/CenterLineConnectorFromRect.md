# CenterLineConnectorFromRect

[![Direct](https://img.shields.io/badge/Direct%20Link-CenterLineConnectorFromRect.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/path/CenterLineConnectorFromRect.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CenterLineConnectorFromRect.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 更新日：2026-04-27
- Excel などからコピー＆ペーストした長方形から罫線（グリッド・枠線）を自動生成
- 線同士の関係に応じて格子化・結合・統合まで実行する Illustrator 用スクリプト

### 主な機能

- 中心線化（任意 ON/OFF）、回転補正、除外条件、格子モード、外枠長方形化に対応
- 線幅は Illustrator の strokeUnits、短辺しきい値は rulerType を参照し、表示値と内部 pt 換算を一元管理
- 線幅の共通化・代表値選択・指定値、印刷用ブラック化、グループ化など仕上げ処理を内蔵
- IIFE 構成で UI 構築／イベント配線／値取得／実行フローを分離し、保守性を確保
- Illustrator DOM の不安定な操作（move/remove/selection 等）は安全ヘルパー経由で実行
- 生成結果は専用作業レイヤーに作成し、ロックレイヤー依存を回避

### 更新履歴

- v1.0.0 (2025-06-12) : 初版作成
- v1.6.0 (2026-04-27) : UI 構造の分離、単位管理の整理、安全操作ヘルパー導入、命名整理、生成専用レイヤー対応
- v1.6.5 (2026-04-27) : UI 構造の改善（事後処理パネル追加、オプション整理、ラベル改善）、生成レイヤー設計の最適化

### スクリプト情報

- バージョン: v1.6.5
