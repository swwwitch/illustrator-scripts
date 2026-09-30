# 〈クリッピングマスクを解除〉を拡張

[![Direct](https://img.shields.io/badge/Direct%20Link-ReleaseClipMask.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/mask/ReleaseClipMask.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReleaseClipMask.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要：

- 選択オブジェクトのクリッピングマスクをモード別に解除できるIllustrator用スクリプト。
- 単純解除、パスのみ削除、画像のみ削除の3つの方法に対応。

<img alt="" src="https://www.dtp-transit.jp/images/ss-548-530-72-20250713-080827.png" width="50%" />

### 主な機能：

- 単純に解除（パスと画像を残す）
- 配置画像を残してパスを削除
- パスを残して配置画像を削除
- パスにK100の塗り（不透明度15%）を適用するオプション付き
- 日本語／英語UI対応
- Q/W/Eキーによるモード切替ショートカット対応

### 処理の流れ：

1. ダイアログでモードとオプションを選択
2. 選択されたモードに従ってマスクを解除
3. 必要に応じてパスに塗りを適用

### note

https://note.com/dtp_tranist/n/nebc832e574f7

### 更新履歴：

- v1.0 (20250606) : 初期バージョン
- v1.1 (20250607) : 安定化と仕様調整
- v1.2 (20250717) : コメント整備
- v1.2.3 (20260928) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.2.4 (20260928) : ボタン行を共通の部品で組むようにした
- v1.2.5 (20260929) : ダイアログの不透明度を98%に変更
- v1.2.6 (20260930) : 文字ツールで文字を選択して実行するとエラーになる不具合を修正
- v1.2.7 (20261001) : 常に左右中央だったボタン行を、ダイアログの内側の幅（左右の余白を除く）が 200px 以内なら中央、それより広ければ右揃えに変更。ウィンドウ・パネルの余白と間隔を共通部品（UIレイアウト）にそろえた
