# RemoveOtherArtboards-v2

[![Direct](https://img.shields.io/badge/Direct%20Link-RemoveOtherArtboards--v2.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/RemoveOtherArtboards-v2.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/RemoveOtherArtboards-v2.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- アクティブなアートボード以外を削除
- アクティブなアートボード外のオブジェクト削除（ガイドを含む）

### 主な機能

- アクティブ以外のアートボードを削除
- アクティブArtboard外のオブジェクト削除
- レイヤー/サブレイヤー/グループのロック・表示・テンプレート状態の一時解除→復元
- アクティブArtboardと重ならないオブジェクトガイドの削除（ルーラーガイドは対象外）
- PageItemのlocked/hidden一時解除→復元

### 処理の流れ

1) アクティブドキュメント・アートボード取得
2) アクティブ以外のアートボードを末尾から削除
3) レイヤー/サブレイヤー/グループ/PageItemの状態を再帰的に収集し一時解除
4) アクティブArtboardと重ならないオブジェクトガイドを削除
5) 反転選択でアクティブArtboard外オブジェクトを削除
6) 状態を復元

### 更新履歴

- v1.0 (202406XX) : 初期バージョン
- v1.1 (20250815) : 説明文強化、ロック復元・ガイド削除ロジックの注記、コメント整理
- v1.2 (20250815) : PageItemのlocked/hidden対応、用語統一、構造化コメント

---

### スクリプト情報

- バージョン: v1.2
