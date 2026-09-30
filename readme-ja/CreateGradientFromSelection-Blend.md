# 配置順の色からグラデーションを生成（ブレンド版、CreateGradientFromSelection に統合）

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

このスクリプトは [CreateGradientFromSelection.jsx](CreateGradientFromSelection.md) に統合しました。

ブレンドは、CreateGradientFromSelection のダイアログで［複製でブレンドを作成］をオンにして使います。

### 更新履歴

- 2026-09-29: CreateGradientFromSelection.jsx に統合して削除
- v1.6
- v1.6.2 : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.6.3 : 一時アクションの読み込み・実行・解除を共通の処理にし、失敗してもアクションセットと一時ファイルが残らないようにした
- v1.6.3 : ボタン行を共通の部品で組むようにした
- v1.6.3 : クリップグループの範囲をマスクで測るようにした（長方形の配置位置に効く）
