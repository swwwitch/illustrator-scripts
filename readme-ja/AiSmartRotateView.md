# AiSmartRotateView

[![Direct](https://img.shields.io/badge/Direct%20Link-AiSmartRotateView.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/misc/AiSmartRotateView.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiSmartRotateView.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

- 開いているドキュメントのアクティブビューの回転角度（表示の回転）を取得
- パレットに回転角度を表示
- 「適用」を押すと、環境設定の「角度の制限」（Shiftキーを押したときの角度）に値を設定

「ビューの回転に連動」をONにすると、たとえば表示を30度回転して作業しているとき、
「角度の制限」にも同じ30度が入り、表示と操作の角度が一致して作業しやすくなります。

パレット（常駐ウィンドウ）化のため、DOMを参照する処理（ビュー回転角度の取得・環境設定の適用）は
ボタンを押すたびにメインエンジンへ BridgeTalk で委譲します。

### スクリプト情報

- バージョン: v1.0.0
- 初回リリース: 20260605
- 最終更新: 20260805
