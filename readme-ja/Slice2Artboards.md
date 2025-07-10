# 画像を分割してアートボード化

[![Direct](https://img.shields.io/badge/Direct%20Link-Slice2Artboards.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/Slice2Artboards.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択した画像やオブジェクトを指定した行数・列数に分割し、各ピースを矩形マスクでクリッピングします。
- さらに各ピースをアートボードとして変換・生成できます。
- 印刷面付け、パズル風レイアウト、複数アートボード化に便利です。

![](https://www.dtp-transit.jp/images/ss-832-724-72-20250711-004617.png)

### 主な機能

- グリッド分割（行数・列数指定）
- オフセットによるサイズ調整
- アスペクト比選択（A4、スクエア、US Letter、16:9、8:9、カスタム）
- アートボードへの自動変換、名前設定、ゼロ埋め
- マージン設定

### 処理の流れ

1. ダイアログで行数・列数、形状（アスペクト比）などを設定
2. OK実行時に分割用マスクを生成
3. 必要に応じてアートボードを追加・リネーム
4. 元画像の削除（オプション）

### 処理の流れ

このスクリプトを実行後、次の流れを想定しています。

1. Illustratorの標準の機能でアートボードを再配置
2. ResizeClipMaskスクリプトでマスクパスの大きさを調整
https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/ResizeClipMask.jsx


### 更新履歴

- v1.0 (20250710): 初期バージョン
- v1.1 (20250710): アートボード変換とオプション追加
- v1.2 (20250710): 形状バリエーション追加、カスタム設定対応
- v1.3 (20250710): 微調整
- v1.4 (20250710): アスペクト比の候補を調整