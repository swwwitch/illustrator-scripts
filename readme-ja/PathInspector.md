# PathInspector

[![Direct](https://img.shields.io/badge/Direct%20Link-PathInspector.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/PathInspector.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PathInspector.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 選択中／全体のパス統計をカウントし、常駐パレットで表示
- 「書き出し」でレポート（テキスト）を書き出し可能
- 常駐パレット化：開いたまま選択を切り替え、［更新］で再集計
- DOM 集計はメインエンジンへ BridgeTalk 委譲（常駐パレットの DOM 切断を回避）

### 主な機能

- パス（オープン／クローズ／アンカー／ハンドル／複合パス／複合シェイプ）を表示
  ※ ガイド（guides=true）はパス統計から除外
- グループ内のパスを再帰的にカウント
- 「書き出し」ボタンで集計結果をテキストファイルとして保存
- ダイアログ位置をセッション中に記憶・復元

### 処理の流れ

- 常駐パレットを表示（多重起動は自動で閉じてから再表示）
- ［更新］押下（または表示直後）にメインエンジンへ集計を委譲
- 戻り値（マーカー方式）を解析し、パス統計パネルを更新
- 書き出しは収集データから生成

---

### スクリプト情報

- バージョン: v1.0.0
- 初回リリース: 20260731
- 最終更新: 20260731
