# 指定文字のベースラインを自動的に天地中央にする

[![Direct](https://img.shields.io/badge/Direct%20Link-AdjustBaselineVerticalCenter.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/AdjustBaselineVerticalCenter.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

選択したテキストフレーム内の指定した文字（1文字以上）を、基準文字に合わせてベースライン（垂直位置）を調整します。

<img alt="" src="https://www.dtp-transit.jp/images/ss-426-384-72-20250713-082008.png" width="50%" />

### 処理の流れ:

1. ダイアログで対象文字と基準文字を指定（対象文字は自動入力、複数ある場合は最頻出記号。手動上書きも可）
2. 複製とアウトライン化で中心Y座標を比較
3. 差分をすべての対象文字に適用

### 対象:

- テキストフレーム（複数選択可、一括適用対応）

### 対象外:

- アウトライン済み、非テキストオブジェクト

### オリジナルアイデア:

Egor Chistyakov https://x.com/tchegr

### オリジナルからの変更点:

- 対象文字は自動入力（複数ある場合には最頻出記号を選択）
- 手動での上書き入力も可能
- 複数のテキストオブジェクトに対しても一括適用可能
- 対象文字に複数文字を同時指定できるよう対応

### 更新履歴:

- v1.0.0(2025-07-04): 初版リリース
- v1.0.6(2025-07-05): 複数の対象文字を指定し、一括で調整可能に対応

### ブログ

https://note.com/dtp_tranist/n/na7a8c907c68c