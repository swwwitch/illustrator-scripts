# 背面に白フチを作る

[![Direct](https://img.shields.io/badge/Direct%20Link-AddOutlineOffsetPath.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/path/AddOutlineOffsetPath.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddOutlineOffsetPath.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択オブジェクトを複製 → 背面配置 → オフセットパス（Live Effect）→ アウトライン → 合体 → 拡張
- 元オブジェクトと結果をグループ化し、Subtract を実行して白で塗りつぶす

### 主な機能

- 複数選択対応
- 単位対応（pt, mm, in, cm など）
- 角の形状（マイター、ラウンド、ベベル）設定可能
- オフセット値のダイアログ入力
- ダイアログの位置調整と透明度設定
- 数値欄は ∧∨ ボタンと ↑↓ キーで次の整数へ増減（1.5→2。Shift で次の10の倍数へ、Option で0.1ずつ）

### 処理の流れ

1) 選択オブジェクトを複製し、背面へ移動
2) オフセットパス（Live Effect）を適用
3) アウトライン化し、合体（Unite）後に拡張（Expand）
4) 元オブジェクトと結果をグループ化し、Subtract を実行
5) 結果を白で塗りつぶす

### クレジット

このスクリプトの一部は、以下のスクリプトを参考にして開発しました。
Outline.jsx (illustrator-outline-script) 作者: Oğuzhan Yıldırım @oguzhanyildirim01
https://github.com/oguzhanyildirim01/illustrator-outline-script/blob/main/Outline.jsx

### 更新履歴

- v1.0.0 (2025-08-13) : 初期バージョン
- v1.1.0 (2025-08-13) : オフセット値の自動計算
- v1.1.2 (2026-09-27) : ↑↓以外のキーを押しても入力値が丸め直される不具合を修正、ドキュメントが開かれていないときのアラートを追加
- v1.2.0 (2026-09-27) : 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.2.1 (2026-09-28) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.2.2 (2026-09-28) : ボタン行を共通の部品で組むようにした
- v1.2.3 (2026-09-29) : ダイアログの不透明度を98%に変更
- v1.2.4 (2026-09-30) : 文字ツールで文字を選択して実行するとエラーになる不具合を修正
- v1.2.5 (2026-09-30) : 初めて開くときにダイアログを右へずらす独自の配置をやめた
- v1.2.6（2026-10-01）常に左右中央だったボタン行を、ダイアログの内側の幅（左右の余白を除く）が 200px 以内なら中央、それより広ければ右揃えに変更。ウィンドウ・パネルの余白と間隔を共通部品（UIレイアウト）にそろえた
- v1.2.7（2026-10-01）ボタン行の下に余白を加え、Illustrator 標準のダイアログに合わせた

### スクリプト情報

- バージョン: v1.2.5
