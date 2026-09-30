# 複数の条件でテキストフレームを一括選択

[![Direct](https://img.shields.io/badge/Direct%20Link-TextSelector.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/TextSelector.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextSelector.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

ドキュメント内のテキストフレームを、複数の条件で一括選択します。

### 主な機能

- 属性で選択: 選択中テキストを基準に、フォントファミリー／＋スタイル／＋サイズ／フォントサイズ／テキストカラー／不透明度で検索
- テキストの種類: すべて／ポイント文字／エリア内文字／パス上文字
- 文字列で選択: 完全一致／部分一致／先頭一致／末尾一致／正規表現
- 選択後の処理: なし／非表示／「_text」レイヤーへ移動／一括編集

### 使い方

1. （属性で選択する場合は）基準にするテキストを選択します。
2. スクリプトを実行します。
3. 条件と選択後の処理を指定して実行します。

### 注意点

- 「_text」レイヤーへ移動する場合、既存レイヤーのロック・可視状態は復元します。
- 一括編集では、書式（`characterAttributes`）を維持したまま内容を置換します。

### 更新履歴

- v1.2.5
- v1.2.7 (2026-09-28) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.2.8 (2026-09-28) : ボタン行を共通の部品で組むようにした
- v1.2.8 (2026-09-28) : キーボードショートカットを共通の部品にした（⌘などを押しているときは反応しない）
- v1.2.9 (2026-09-29) : ダイアログの不透明度を98%に変更
- v1.2.10 (2026-09-30) : 文字ツールで文字を選択して実行するとエラーになる不具合を修正
- v1.2.11 (2026-09-30) : 右側のボタンだけのボタン行を左右中央に並べるようにした
- v1.2.12（2026-10-01）右側のボタンだけの行は、ダイアログの内側の幅（左右の余白を除く）が 200px 以内なら中央、それより広ければ右揃えに変更
- v1.2.13（2026-10-01）ウィンドウ・パネルの余白と間隔を共通部品（UIレイアウト）にそろえた
