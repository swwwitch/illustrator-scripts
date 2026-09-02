# TypeBasicsPanel

[![Direct](https://img.shields.io/badge/Direct%20Link-TypeBasicsPanel.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/TypeBasicsPanel.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TypeBasicsPanel.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 概要

選択したテキストの基本的な文字組み設定（フォントサイズと行送り・自動カーニング・
プロポーショナルメトリクス・文字ツメ・トラッキング）だけをまとめて行う常駐パレットスクリプトです。
UnifiedTypePanel.jsx からこれら5機能を抜き出して1画面に集約しています。

- フォントサイズと行送り：サイズ・実質（pt）・行送り（%）の3入力を1パネルに。
  % を段落の「自動行送り量」に設定し常に自動行送りにする（Illustrator 上は常に「自動」表示、
  行送りはサイズに自動追従）。実質＝サイズ×%。実質欄に値を入れると % を逆算して設定。
- 自動カーニング：和文等幅／0／メトリクス／オプティカル（「メトリクス」のみプロポーショナルメトリクスON）
- 文字ツメ：0〜100%。入力欄とスライダー（Shift で 10% 刻み）
- トラッキング：-100〜500。入力欄とスライダー（Shift で 10 刻み）
- プロポーショナルメトリクス：チェックボックスで単独にオン／オフ（段落単位で適用）
- ラジオや入力を操作すると、その場で選択中のテキストへ即時適用する
- 選択は単体のテキストフレームだけでなく、グループ内のテキストやテキスト編集モードでの範囲選択にも対応
- パレットにフォーカスが戻るたび、または「再読み込み」で選択の現在値を読み取って UI に反映する
- 常駐エンジン（#targetengine）でパレット表示。常駐エンジンの app は
  パレット表示中に DOM 接続を失うため、DOM 処理はメインエンジンへ
  BridgeTalk で都度委譲する（コードは encodeURIComponent で包んで送信）

### スクリプト情報

- バージョン: v1.0.3

### note

- [フォントサイズを変更しても行送りが変わらない問題｜DTP Transit 別館](https://note.com/dtp_tranist/n/n29e7115b5e70)
- 上位版：[統合文字組みパレット｜DTP Transit 別館](https://note.com/dtp_tranist/n/n4e2b79cf2891)

