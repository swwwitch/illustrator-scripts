# KPTSketchy

[![Direct](https://img.shields.io/badge/Direct%20Link-KPTSketchy.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/KPTSketchy.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/KPTSketchy.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択したオブジェクトにランダムな変形を加え、手書き・スケッチ風の見た目に整えます。
- 塗り／線の分離、角丸、パスのオフセット、ラフ効果（ギザギザ／歪曲）、グループ化に対応します。
- 変形はすべてライブ効果として適用するため、あとから編集・解除できます。
- 常駐パレットとして表示され、開いたままドキュメントを操作できます。

### 使い方

1. スクリプトを実行するとパレットが表示されます。
2. オブジェクトを選択し、各パネルで効果を調整します（変更するたびに結果が作り直されます）。
3. ［再計算］で乱数を振り直します。気に入ったらそのままにしておくだけで確定です。
4. パレットから離れる、または閉じる（× / Esc）と、その時点の結果が確定します。

### 仕様

- 角丸・オフセット・移動は、環境設定の単位（rulerType）で入力します。
- 適用順は「角丸 → パスのオフセット → 変形 → ラフ効果（歪曲 → ギザギザ）」です。
- オフセットは負方向 → 正方向の順に 2 回適用します。
- 塗りと線の両方を持つオブジェクトのみ分離できます（塗りのみ／線のみは分離しません）。
- グループは「グループ内を個別に処理」ON のとき、中の各オブジェクトへ展開します。
- クリッピングパスは処理対象から除外します。
- 値の変更・再計算では、直前の結果を app.undo() で取り消してから適用し直します。
- パレットからフォーカスが外れた時点で結果を確定し、以降は取り消しません
  （ユーザーがドキュメント側で行った操作を誤って取り消さないため）。

### 実装メモ

- 常駐パレットの app はドキュメントへの接続を失うため、DOM を触る処理は
  すべて worker 関数にまとめ、BridgeTalk でメインエンジンへ委譲します。
- worker 関数は toString() で連結し、encodeURIComponent したうえで
  eval(decodeURIComponent(...)) の形で送信します。
- worker 関数内では行コメントを使わず、必ずセミコロンで文を終えます。

---

### スクリプト情報

- バージョン: v1.2.0
- 初回リリース: 2026-04-14
- 最終更新: 2026-07-22
