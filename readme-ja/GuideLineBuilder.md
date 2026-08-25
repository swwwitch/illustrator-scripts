# GuideLineBuilder

[![Direct](https://img.shields.io/badge/Direct%20Link-GuideLineBuilder.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/GuideLineBuilder.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GuideLineBuilder.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択オブジェクト内のパス（グループ／複合パスを含む）から隣接するアンカーポイントのペアを取り、補助線として「直線を描画範囲いっぱいに延長した線」を描画します。

### 主な機能

- ［直線］ONで直線セグメントのみ、OFFで曲線セグメントのみを対象
- ［円弧から円］: Bezier曲線セグメントから円を推定して作成。正確な円弧でない場合は［円弧オプション］で 無視／直線（弦）／直線（延長）を選択
- 線幅は環境設定の「線」（`strokeUnits`）の単位に追従（既定は 0.1mm 相当）
- プレビューONで、ダイアログを閉じる前に一時レイヤーへ描画し、終了時に自動で消去

### 使い方

1. 対象のパスを選択します。
2. スクリプトを実行します。
3. オプションを指定して実行します。

### 注意点

- 選択がアクティブアートボードと交差しない場合は、選択の外接幅A・高さBを基準に中心へ矩形（幅A×4、高さB×4）を仮想し、その中で延長線を描画します。
- ［別レイヤーに］ONのときは `_construction_guide` レイヤーへ出力します。選択がそのレイヤー上にある場合は、既存レイヤーを `_construction_guide_backup...` に退避して新しいレイヤーを作成します（元オブジェクトは削除しません）。

### 更新履歴

- v1.2 (2026-03-12)
