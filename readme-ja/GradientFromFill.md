# 元の塗り色を始点にした線形グラデーションを作成

[![Direct](https://img.shields.io/badge/Direct%20Link-GradientFromFill.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/GradientFromFill.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GradientFromFill.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択した塗りオブジェクトに対して、元の塗り色を始点にした線形グラデーションを作成します。

### 主な機能

- 単色オブジェクトはその塗り色を始点に使用
- 塗りがグラデーションの単一オブジェクトでは、黒・白・透明を除いたストップから始点色を選択
- 終点カラーは黒／白／透明／補色／淡色から選択
- 角度は 0 / 30 / 45 / 60 / 90 度から選択
- セパレートグラデーション、反転、プレビューに対応

### 使い方

1. 対象のオブジェクトを選択します。
2. スクリプトを実行します。
3. 始点カラー・終点カラー・角度を指定して［OK］をクリックします。

### 注意点

- 複数オブジェクトを選択している場合は、始点カラーパネル全体が無効になります。
- キャンセル時は元の塗り色へ戻します。プレビュー中に例外が発生した場合も、元の塗り色と選択状態を復元します。
- 複合パス、グループ内の再帰処理、クリッピンググループ内のオブジェクトにも対応します。

### 更新履歴

- v1.1.0
- v1.1.2 : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.1.3 : ボタン行を共通の部品で組むようにした
- v1.1.3 : キーボードショートカットを共通の部品にした（⌘などを押しているときは反応しない）
- v1.1.4 : ダイアログの不透明度を98%に変更
- v1.1.5 : 文字ツールで文字を選択して実行するとエラーになる不具合を修正
- v1.1.7（2026-10-01）ボタン行の下に余白を加え、Illustrator 標準のダイアログに合わせた
- v1.1.6（2026-10-01）常に左右中央だったボタン行を、ダイアログの内側の幅（左右の余白を除く）が 200px 以内なら中央、それより広ければ右揃えに変更。ウィンドウ・パネルの余白と間隔を共通部品（UIレイアウト）にそろえた
