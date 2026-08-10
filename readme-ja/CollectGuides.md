# CollectGuides

[![Direct](https://img.shields.io/badge/Direct%20Link-CollectGuides.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/guide/CollectGuides.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CollectGuides.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 複数のレイヤー／サブレイヤーに散在するガイドを、1 つのレイヤー（既定は `// guide`）へ集約します。
- 非表示・ロックされたレイヤーやガイドも対象にし、処理中は一時的に解除して、終了後に元の状態へ戻します。

### 主な機能

- 全レイヤー（サブレイヤー含む）を再帰的に走査してガイドを移動
- レイヤー／アイテムの locked・visible・hidden を退避して復元
- ターゲットレイヤー名を定数 `TARGET_GUIDE_LAYER_NAME` で変更可能
- 実行後のガイド表示／ロック状態を任意の状態へ戻すオプション
- 大規模ドキュメント向けに、対象レイヤーだけを走査する最適化と Redraw 抑制

### 使い方

1. 対象のドキュメントを開く
2. スクリプトを実行する

### 注意点

- 実行中は「ガイドを表示」「ガイドのロックを解除」を一時的に実行します。
- ターゲットレイヤーが既にある場合はそれを再利用します。

### 更新履歴

- v1.0 (20250816): 初版。ガイド集約、再帰走査、ロック／可視状態の復元
- v1.1 (20250816): ガイドを含むレイヤーだけに絞り込む対象限定スキャンを追加
- v1.2 (20250816): 実行後のガイド表示／ロック状態の復帰設定を追加
- v1.3 (20250816): ターゲットレイヤー名を定数で切り替え可能に
- v1.4 (20250816): Redraw 抑制オプションを追加
