# パスに沿ってオブジェクトを等間隔に配置

[![Direct](https://img.shields.io/badge/Direct%20Link-ArrangeObjectsAlongPath.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/ArrangeObjectsAlongPath.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArrangeObjectsAlongPath.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

複数のオブジェクトを、選択範囲内の1本のパスに沿って等間隔に自動配置します。

### 主な機能

- 基準のパスは「自動（面積最大）／最前面／最背面」から指定
- 複製パネルで、配置しながら複製する数を指定

### 使い方

1. 配置したいオブジェクトと、基準にするパスをまとめて選択します。
2. スクリプトを実行します。
3. 基準パスの決め方と複製数を指定して実行します。

### 注意点

- 「自動（面積最大）」は、開パスの場合はバウンディングボックスの面積で判定します。
- 複製パネルは、選択がちょうど2つのときだけ既定で有効になり、複製数の初期値は2です。

### 更新履歴

- v1.6.7（2026-10-01）ボタン行の下に余白を加え、Illustrator 標準のダイアログに合わせた
- v1.6.6（2026-10-01）ウィンドウ・パネルの余白と間隔を共通部品（UIレイアウト）にそろえた
- v1.6.5 (2026-09-30) : 初めて開くときにダイアログを右へずらす独自の配置をやめた
- v1.6.4 (2026-09-30) : 文字ツールで文字を選択して実行するとエラーになる不具合を修正
- v1.6.3 (2026-09-29) : ダイアログの不透明度を98%に変更
- v1.6.2 (2026-09-28) : ボタン行を共通の部品で組むようにした
- v1.6.1 (2026-09-28) : ダイアログを前回閉じた位置で開き、選択中のオブジェクトに重なるときは左右にずらすようにした。不透明度を97%にそろえた
- v1.6.0 (2026-09-27) : 数値欄にステップボタン（∧∨）を追加。↑↓キーもステップボタンと同じ処理で増減するように変更（次の整数へ、shift＋で次の10の倍数へ）
- v1.5.0 (2026-03-03)
